{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/11/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : COTISCONTRAT (COTISCONTRAT)
Mots clefs ... : TOM;COTISCONTRAT
*****************************************************************}
{
PT1 GGU 28/11/2007 FQ 14983 vérifier la référence contrat avec le code institution
}

Unit UTomCOTISCONTRAT ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList, 
     HDB,
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
  TOM_COTISCONTRAT = Class (TOM)
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
  public
    procedure OnElipsisRefContratClick(Sender : TObject);
  end ;
                                  
Implementation

uses
  lookup;

procedure TOM_COTISCONTRAT.OnNewRecord ;
begin
  Inherited ;
  SetField('POC_NATURERUB', 'COT');
end ;

procedure TOM_COTISCONTRAT.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_COTISCONTRAT.OnUpdateRecord ;
var
  RechQuery : TQuery;  //PT1
  Procedure FieldFocus(FieldName : String);
  begin
    if Assigned(GetControl(FieldName)) then SetFocusControl(FieldName);
  end;
  Procedure ControleChampsNonVide(FieldName : String; MsgLibelleChamps : String = '');
  begin
    if (GetField(FieldName) = '') then
    begin
      LastError := 1;
      if MsgLibelleChamps = '' then MsgLibelleChamps := FieldName;
      LastErrorMsg := 'Vous devez renseigner le champ '+MsgLibelleChamps+'.';
      FieldFocus(FieldName);
    end;
  end;
begin
  if (GetField('POC_INSTITUTION') <> '') and (not IsNumeric(GetField('POC_INSTITUTION'))) then
  begin
    SetField('POC_INSTITUTION', '');
  end;
  ControleChampsNonVide('POC_INSTITUTION', 'institution');
  ControleChampsNonVide('POC_REFERCONTRAT', 'référence du contrat');
  ControleChampsNonVide('POC_PREDEFINI', 'prédéfini');
  ControleChampsNonVide('POC_NODOSSIER', 'numéro de dossier');
  ControleChampsNonVide('POC_NATURERUB', 'nature de la rubrique');
  ControleChampsNonVide('POC_RUBRIQUE', 'rubrique');
  //debut PT1
  RechQuery := OpenSQL('SELECT count(*) FROM CONTRATPREV WHERE POP_REFERCONTRAT = "'+GetField('POC_REFERCONTRAT')+'" AND POP_INSTITUTION = "'+GetField('POC_INSTITUTION')+'" ', True);
  if (RechQuery.Eof) or (RechQuery.Fields[0].AsInteger <= 0) then
  begin
    SetField('POC_REFERCONTRAT', '');
    FieldFocus('POC_REFERCONTRAT');
    LastError := 1;
    LastErrorMsg := 'La référence du contrat de prévoyance ne concorde pas avec l''institution choisie.';
  end;
  Ferme(RechQuery);
  //Fin PT1
  Inherited ;
end ;

procedure TOM_COTISCONTRAT.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_COTISCONTRAT.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_COTISCONTRAT.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_COTISCONTRAT.OnChangeField ( F: TField ) ;
var
  Predef, NoDos : String;
  RechQuery : TQuery;
begin
  Inherited ;
  if F.FieldName = 'POC_INSTITUTION' then
  begin
    if (GetField('POC_INSTITUTION') <> '') and (not IsNumeric(GetField('POC_INSTITUTION'))) then
    begin
      SetField('POC_INSTITUTION', '');
    end;
    if Assigned(GetControl('POC_REFERCONTRAT')) then
    begin
      if GetField('POC_INSTITUTION') <> '' then
      begin
        (GetControl('POC_REFERCONTRAT') as {$IFNDEF EAGLCLIENT}THDBEdit{$else}THEdit{$ENDIF}).Plus := ' AND POP_INSTITUTION ="'+GetField('POC_INSTITUTION')+'"';
      end else begin
        (GetControl('POC_REFERCONTRAT') as {$IFNDEF EAGLCLIENT}THDBEdit{$else}THEdit{$ENDIF}).Plus := '';
      end;
    end;
  end;
  if F.FieldName = 'POC_RUBRIQUE' then
  begin
    { On cherche dans la base le predefini et le nodossier de la cotisation }
    if GetField('POC_RUBRIQUE') <> '' then
    begin
      Predef := '';
      NoDos := '';
      RechQuery := OpenSQL('SELECT PCT_PREDEFINI, PCT_NODOSSIER FROM COTISATION WHERE ##PCT_PREDEFINI## AND PCT_NATURERUB = "COT" AND PCT_RUBRIQUE = "'+GetField('POC_RUBRIQUE')+'" ', True);
      if not RechQuery.Eof then
      begin
        Predef := RechQuery.FindField('PCT_PREDEFINI').AsString;
        NoDos  := RechQuery.FindField('PCT_NODOSSIER').AsString;
      end else begin
        PGIError('La cotisation saisie n''est pas valide.');
        SetField('POC_RUBRIQUE', '');
      end;
      Ferme(RechQuery);
      SetField('POC_PREDEFINI', Predef);
      SetField('POC_NODOSSIER', NoDos);
    end;
  end;
end ;

procedure TOM_COTISCONTRAT.OnArgument ( S: String ) ;
begin
  Inherited ;
  if Assigned(GetControl('POC_REFERCONTRAT')) then
    (GetControl('POC_REFERCONTRAT') as {$IFNDEF EAGLCLIENT}THDBEdit{$else}THEdit{$ENDIF}).OnElipsisClick := OnElipsisRefContratClick;
end ;

procedure TOM_COTISCONTRAT.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_COTISCONTRAT.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_COTISCONTRAT.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_COTISCONTRAT.OnElipsisRefContratClick(Sender: TObject);
var
  stWhere : String;
begin
  stWhere := (GetControl('POC_REFERCONTRAT') as {$IFNDEF EAGLCLIENT}THDBEdit{$else}THEdit{$ENDIF}).plus;
  LookupList(GetControl('POC_REFERCONTRAT'), 'Contrats de prévoyances', 'CONTRATPREV', 'POP_REFERCONTRAT',
             'POP_TYPECONTRAT, POP_NOMPOP', stWhere,
             'POP_REFERCONTRAT', True, 8, '', tlDefault, 50);
end;

Initialization
  registerclasses ( [ TOM_COTISCONTRAT ] ) ; 
end.
