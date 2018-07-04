{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 26/01/2005
Modifié le ... :   /  /
Description .. : Source TOT de la TABLETTE : TTCATJALBUD (TTCATJALBUD)
Mots clefs ... : TOT;TTCATJALBUD
*****************************************************************}
Unit TTCATJALBUD_TOT ;

Interface

Uses StdCtrls, Controls, Classes, UProcGen,
{$IFDEF EAGLCLIENT}
     UTob, eTablette,
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     db, Tablette,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, HPanel,
     UTOT;

procedure ParamStructureBudget(PRien : THPanel);

Type
  TOT_TTCATJALBUD = Class ( TOT )
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnUpdateRecord           ; override ;
    procedure OnAfterUpdateRecord      ; override ;
    procedure OnClose                  ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnComplement(Prefixe, Tipe, Code : string) ; override;
    function  ChercheNewCode : String;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ELSE}
  Ent1,     // MaxCatBud
  {$ENDIF MODENT1}
  ParaSpBu; // ParamSousPlanBudget


procedure ParamStructureBudget(PRien : THPanel);
begin
  ParamTable('TTCATJALBUD', taCreat, 7577200, PRien, 3, 'Catégories de budget');
end;

procedure TOT_TTCATJALBUD.OnNewRecord () ;
begin
  Inherited ;
  {$IFDEF EAGLCLIENT}
  if DS.Detail.Count >=MaxCatBud then begin
  {$ELSE}
  if DS.RecordCount >=MaxCatBud then begin
  {$ENDIF EAGLCLIENT}
    HShowMessage('8;Catégories de budget;Nombre maximum d''enregistrement atteint. Vous ne pouvez plus créer de nouvel enregistrement.;W;O;O;O;', '', '');
    LastError := 1;
  end;

  SetField('CC_CODE', ChercheNewCode);
end ;

procedure TOT_TTCATJALBUD.OnDeleteRecord () ;
begin
  Inherited ;

  if HShowMessage('7;Catégories de budget;Les budgets rattachés à ce code vont être mis à jour. Désirez-vous continuer ?;Q;YNC;N;N;', '', '')<>mrYes then begin
    LastError := 1;
    Exit;
  end;
  ExecuteSql('UPDATE BUDJAL SET BJ_CATEGORIE="" WHERE BJ_CATEGORIE="'+GetField('CC_CODE')+'"');
  ExecuteSql('DELETE FROM CROISCPT WHERE CX_TYPE="BUD" AND CX_JAL="'+GetField('CC_CODE')+'" AND CX_CATEGORIE="'+GetField('CC_CODE')+'"');
end ;

procedure TOT_TTCATJALBUD.OnUpdateRecord () ;
begin
  Inherited ;

  if GetField('CC_ABREGE')='' then begin
    HShowMessage('9;;Vous devez paramétrer les sous-plans des catégories.;W;O;O;O;', '', '');
    LastError := 1;
  end;
end ;

procedure TOT_TTCATJALBUD.OnAfterUpdateRecord () ;
begin
  Inherited ;
  AvertirTable('TTCATJALBUD');
  MajInfoTablette('TTCATJALBUD');
end ;

procedure TOT_TTCATJALBUD.OnClose () ;
begin
  Inherited ;
end ;

procedure TOT_TTCATJALBUD.OnArgument (S : String ) ;
begin
  Inherited ;
  OkComplement := True;
  bComplement.Hint := TraduireMemoire('Paramétrage des sous-plans budgétaires');
  bComplement.GlobalIndexImage := 'Z0008_S16G1';

  {JP 15/10/07 : FQ 21645 : Ecran est disponible pour les tablettes, plus besoin de le récupérer}
  {$IFDEF EAGLCLIENT}
  TFTablette(Ecran).FListe.ColWidths[2] := -1;
  TFTablette(Ecran).FListe.ColEditables[0] := False;
  {$ELSE}
  TFTablette(Ecran).Fliste.Columns[2].Visible := False;
  TFTablette(Ecran).Fliste.Columns[2].ReadOnly := True;
  {$ENDIF EAGLCLIENT}
end ;

procedure TOT_TTCATJALBUD.OnComplement(Prefixe, Tipe, Code: string);
var
  Axe,SousPlan : String;
begin
  inherited;

  if (Code = '') or (GetField('CC_LIBELLE') = '') then Exit;
  Axe := GetField('CC_ABREGE');
  SousPlan := GetField('CC_LIBRE');
  ParamSousPlanBudget(Code, GetField('CC_LIBELLE'), Axe, SousPlan);

  if (Axe <> GetField('CC_ABREGE')) or (SousPlan<> GetField('CC_LIBRE')) then begin
    if DS.State=dsBrowse then DS.Edit;
    SetField('CC_ABREGE', Axe);
    SetField('CC_LIBRE', SousPlan);
  end;
end;

function TOT_TTCATJALBUD.ChercheNewCode: String;
var
  QLoc : TQuery;
  i : Integer;
begin
  {$IFDEF EAGLCLIENT}
  if DS.Detail.Count=0 then begin
  {$ELSE}
  if DS.RecordCount=0 then begin
  {$ENDIF EAGLCLIENT}
    Result:='S01';
    Exit;
  end;

  QLoc := OpenSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="CJB" ORDER BY CC_CODE', True);
  i := 1;
  While not QLoc.Eof do begin
    if Copy(QLoc.Fields[0].AsString,3,1)<>IntToStr(i) then Break;
    Inc(i);
    QLoc.Next;
  end;
  Ferme(QLoc);
  if i<=9 then Result := 'S0'+IntToStr(i)
          else Result := 'S' +IntToStr(i);
end;

Initialization
  registerclasses ( [ TOT_TTCATJALBUD ] ) ; 
end.
