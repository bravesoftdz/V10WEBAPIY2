{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 16/10/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTAPPELINTCOMPL ()
Mots clefs ... : TOF;BTAPPELINTCOMPL
*****************************************************************}
Unit BTAPPELINTCOMPL_TOF ;
 
Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
		 AglInit,	
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     LookUp,
     UtilGc,
     DicoBTP,
     HRichEdt, HRichOLE ;

Type
  TOF_BTAPPELINTCOMPL = Class (TOF)
  private
  	TOBInfoCOmpl,TOBInfosSup : TOB;
    TT : TOB;
    procedure  StoreChamps;
    procedure  RecupChamps;
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

const  TexteMsgAffaire: array [1..1] of string = (
    {1}'saisi n''existe pas (table libre)'
    ) ;

Implementation

procedure TOF_BTAPPELINTCOMPL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPELINTCOMPL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPELINTCOMPL.OnUpdate ;
begin
  Inherited ;
  if GetControltext ('AFF_LIBREAFF1') <> '' then
  if not LookupValueExist (GetControl ('AFF_LIBREAFF1') ) then
  begin
    LastError := 1;
    LastErrorMsg := TraduitGa (THLabel (GetControl ('TAFF_LIBREAFF1') ) .Caption + ' ' + TexteMsgAffaire [LastError] ) ;
    SetFocusControl ('AFF_LIBREAFF1') ;
    Exit;
  end;
  if GetControlText ('AFF_LIBREAFF2') <> '' then
  if not LookupValueExist (GetControl ('AFF_LIBREAFF2') ) then
  begin
    LastError := 1;
    LastErrorMsg := TraduitGa (THLabel (GetControl ('TAFF_LIBREAFF2') ) .Caption + ' ' + TexteMsgAffaire [LastError] ) ;
    SetFocusControl ('AFF_LIBREAFF2') ;
    Exit;
  end;
  if GetCOntroltext ('AFF_LIBREAFF3') <> '' then
  if not LookupValueExist (GetControl ('AFF_LIBREAFF3') ) then
  begin
    LastError := 1;
    LastErrorMsg := TraduitGa (THLabel (GetControl ('TAFF_LIBREAFF3') ) .Caption + ' ' + TexteMsgAffaire [LastError] ) ;
    SetFocusControl ('AFF_LIBREAFF3') ;
    Exit;
  end;
  //
  StoreChamps;
  TheTOB := TOBInfoCOmpl;
  //
end ;

procedure TOF_BTAPPELINTCOMPL.OnLoad ;
begin
  Inherited ;
  RecupChamps;
end ;

procedure TOF_BTAPPELINTCOMPL.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBInfoCOmpl := LaTOb;
  TT := TOB(LaTOB.Data);
  TOBInfosSup := TOB(TT.Data);
  // Paramétrage des libellés des tables libres en gescom
  GCMAJChampLibre (TForm (Ecran) , False, 'COMBO', 'AFF_LIBREAFF', 10, '') ;
  GCMAJChampLibre (TForm (Ecran) , False, 'EDIT', 'AFF_VALLIBRE', 3, '') ;
  GCMAJChampLibre (TForm (Ecran) , False, 'EDIT', 'AFF_DATELIBRE', 3, '') ;
  GCMAJChampLibre (TForm (Ecran) , False, 'EDIT', 'AFF_CHARLIBRE', 3, '') ;
  GCMAJChampLibre (TForm (Ecran) , False, 'EDIT', 'AFF_RESSOURCE', 3, '') ;
  GCMAJChampLibre (TForm (Ecran) , False, 'BOOL', 'AFF_BOOLLIBRE', 3, '') ;

end ;

procedure TOF_BTAPPELINTCOMPL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPELINTCOMPL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPELINTCOMPL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPELINTCOMPL.RecupChamps;
begin
  TOBInfoCOmpl.PutEcran(ecran);
  TT.PutEcran(ecran);
end;

procedure TOF_BTAPPELINTCOMPL.StoreChamps;
begin
  TOBInfoCOmpl.GetEcran(ecran);
  TT.GetEcran(ecran);
end;

Initialization
  registerclasses ( [ TOF_BTAPPELINTCOMPL ] ) ;
end.
