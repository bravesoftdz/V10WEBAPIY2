{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/10/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PUBLICOTIS (PUBLICOTIS)
Mots clefs ... : TOM;PUBLICOTIS
*****************************************************************}
{
PT1    : 09/11/2007 VG V_80 Prise en compte des FQ N°13646
}
Unit UTOMPUBLICOTIS;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     HDB,
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
     UTob,
     LookUp,
     PgOutils,
     PgOutils2;

Type
  TOM_PUBLICOTIS = Class (TOM)
    public
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
    private
{$IFNDEF EAGLCLIENT}
    Rubrique : THDBEdit;
{$ELSE}
    Rubrique : THEdit;
{$ENDIF}
    CEG, DOS, STD : boolean;

    end;

Implementation

procedure TOM_PUBLICOTIS.OnNewRecord;
begin
Inherited;
//PT1
SetField ('PUO_PREDEFINI', 'DOS');
SetField ('PUO_NODOSSIER', PgRendNoDossier);
//FIN PT1
end;

procedure TOM_PUBLICOTIS.OnDeleteRecord;
begin
Inherited;
end;

procedure TOM_PUBLICOTIS.OnUpdateRecord;
var
Predef : string;
begin
Inherited;
//PT1
if (DS.State=dsinsert) then
   begin
   if (GetField ('PUO_PREDEFINI')<>'DOS') then
       SetField ('PUO_NODOSSIER', '000000')
   else
       SetField ('PUO_NODOSSIER', PgRendNoDossier);
  end;

Predef:= GetField('PUO_PREDEFINI');
if (Predef<>'CEG') and (Predef<>'DOS') and (Predef<>'STD') then
   begin
   LastError:= 1;
   PgiBox ('Vous devez renseigner le champ prédéfini.', Ecran.caption);
   SetFocusControl ('PUO_PREDEFINI');
   end;

if (GetField ('PUO_NATURERUB')='') then
   begin
   LastError:= 1;
   PGIBox('La nature est obligatoire', Ecran.Caption);
   SetFocusControl ('PUO_NATURERUB');
   Exit;
   end;

if (GetField ('PUO_RUBRIQUE')='') then
   begin
   LastError:= 1;
   PGIBox('La rubrique est obligatoire', Ecran.Caption);
   SetFocusControl ('PUO_RUBRIQUE');
   Exit;
   end;
//FIN PT1
end;

procedure TOM_PUBLICOTIS.OnAfterUpdateRecord;
begin
Inherited;
end;

procedure TOM_PUBLICOTIS.OnAfterDeleteRecord;
begin
Inherited;
end;

procedure TOM_PUBLICOTIS.OnLoadRecord;
begin
Inherited;
//PT1
AccesPredefini ('TOUS', CEG, STD, DOS);
if (Getfield ('PUO_PREDEFINI')='CEG') then
   SetControlEnabled ('BDelete', CEG)
else
if (Getfield ('PUO_PREDEFINI')='STD') then
   SetControlEnabled ('BDelete', STD)
else
if (Getfield ('PUO_PREDEFINI')='DOS') then
   SetControlEnabled ('BDelete', DOS);

SetControlEnabled ('PUO_PREDEFINI', False);

if (DS.State in [dsInsert]) then
   SetControlEnabled ('PUO_PREDEFINI', True);
//FIN PT1
end;

procedure TOM_PUBLICOTIS.OnChangeField ( F: TField );
var
Pred : string;
begin
Inherited;
//PT1
if (F.FieldName='PUO_PREDEFINI') and (DS.State=dsinsert) then
   begin
   Pred:= GetField ('PUO_PREDEFINI');
   if (Pred='') then
      exit;
   AccesPredefini ('TOUS', CEG, STD, DOS);
   if (Pred='CEG') and (CEG=FALSE) then
      begin
      PGIBox ('Vous ne pouvez pas créer d''affectation prédéfinie CEGID',
              'Accès refusé');
      Pred:= 'DOS';
      SetControlProperty ('PUO_PREDEFINI', 'Value', Pred);
      end;
    if (Pred='STD') and (STD=FALSE) then
       begin
       PGIBox ('Vous ne pouvez pas créer d''affectation prédéfinie Standard',
               'Accès refusé');
       Pred:= 'DOS';
       SetControlProperty ('PUO_PREDEFINI', 'Value', Pred);
       end;
   if (Pred<>GetField ('PUO_PREDEFINI')) then
      SetField ('PUO_PREDEFINI', pred);
   end;
//FIN PT1

if (F.Fieldname = 'PUO_NATURERUB') then
   begin
   if (GetField ('PUO_NATURERUB')='REM') then
      begin
      Rubrique.DataType:= 'PGREMUNERATION';
//PT1
      THValComboBox (GetControl ('PUO_CHAMPRETENU')).Plus:= ' AND CO_ABREGE="R" AND'+
                                                            ' CO_CODE<>"09" AND'+
                                                            ' CO_CODE<>"10"';
//FIN PT1
      SetControlEnabled ('PUO_CHAMPRETENU', True);
      end
   else
   if (GetField ('PUO_NATURERUB') = 'COT') then
      begin
      Rubrique.DataType:= 'PGCOTISATION';
      THValComboBox (GetControl ('PUO_CHAMPRETENU')).Plus:= ' AND CO_ABREGE="C"';//PT1
      SetControlEnabled ('PUO_CHAMPRETENU', True);
{PT1
      if (GetField ('PUO_CHAMPRETENU')='COE') then
         SetField ('PUO_CHAMPRETENU', '');
}
      end
   else
      begin
      Rubrique.DataType:= 'PGCUMULPAIE';
      SetControlText ('PUO_CHAMPRETENU', '');
      SetControlEnabled ('PUO_CHAMPRETENU', False);
      end;
   end;

{PT1
if (F.Fieldname = 'PUO_CHAMPRETENU') then
   begin
   if (GetField ('PUO_CHAMPRETENU')='COE') then
      begin
      PGIBox ('Valeur interdite', Ecran.Caption);
      SetField ('PUO_CHAMPRETENU', '');
      end;
   end;
}

if (Ds.State in [DsBrowse]) then
   SetControlEnabled ('PUO_PREDEFINI', False);
//FIN PT1
end;

procedure TOM_PUBLICOTIS.OnArgument ( S: String );
begin
Inherited;
{$IFNDEF EAGLCLIENT}
Rubrique:= THDBEdit (GetControl ('PUO_RUBRIQUE'));
{$ELSE}
Rubrique:= THEdit (GetControl ('PUO_RUBRIQUE'));
{$ENDIF}
if (S = 'REM') then
   begin
   Rubrique.DataType:= 'PGREMUNERATION';
//PT1
   THValComboBox (GetControl ('PUO_CHAMPRETENU')).Plus:= ' AND CO_ABREGE="R" AND'+
                                                         ' CO_CODE<>"09" AND'+
                                                         ' CO_CODE<>"10"';
//FIN PT1
   end
else
if (S = 'COT') then
   begin
   Rubrique.DataType:= 'PGCOTISATION';
   THValComboBox (GetControl ('PUO_CHAMPRETENU')).Plus:= ' AND CO_ABREGE="C"';//PT1
   end
else
   begin
   Rubrique.DataType:= 'PGCUMULPAIE';
   end;
end;

procedure TOM_PUBLICOTIS.OnClose;
begin
Inherited;
THValComboBox (GetControl ('PUO_CHAMPRETENU')).Plus:= '';  //PT1
end;

procedure TOM_PUBLICOTIS.OnCancelRecord;
begin
Inherited;
end;


Initialization
registerclasses ([TOM_PUBLICOTIS]);
end.
