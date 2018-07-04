{***********UNITE*************************************************
Auteur  ...... : Laroche Vincent
Créé le ...... : 30/08/2004
Modifié le ... :   /  /
Description .. : Source TOT de la TABLE : TTVENTILTYPE (TTVENTILTYPE)
Mots clefs ... : TOF;TTVENTILTYPE
*****************************************************************}
Unit UTOTVENTILTYPE;

Interface

uses
{$IFDEF EAGLCLIENT}
  eTablette, // ParamTable
{$ELSE}
  Tablette, // ParamTable
{$ENDIF}
  Classes, // TMemoryStream, HexToBin
  UTot,
  HPanel,  // THPanel
  HCtrls,  // ExecuteSQL,
  HEnt1,   // taModif
  HTB97,   // TToolbarButton97
  Ent1, //VH^
  Ventil,   // ParamVentil
  CPVENTILTYPECROIS_TOF //ParamVentilCroisaxe
  ;

procedure ParamVentilType(PP : THPanel = nil);

Type
  TOT_TTVENTILTYPE = Class ( TOT )
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnUpdateRecord           ; override ;
    procedure OnAfterUpdateRecord      ; override ;
    procedure OnClose                  ; override ;
    procedure OnArgument  (S : String ); override ;
    procedure OnComplement(Prefixe, Tipe, Code : string); override ;
  private
    bCreate : Boolean;
  end ;

implementation

procedure ParamVentilType(PP : THPanel = nil);
begin
  ParamTable('ttVentilType', taCreat,1460000, PP, 3, TraduireMemoire('Ventilations types')) ; //XVI 24/02/2005
end;

procedure TOT_TTVENTILTYPE.OnNewRecord ;
begin
  Inherited ;
  bCreate := True;
end ;

procedure TOT_TTVENTILTYPE.OnDeleteRecord ;
begin
  Inherited ;
  ExecuteSql('DELETE FROM VENTIL WHERE V_NATURE LIKE "TY%" AND V_COMPTE="'+GetField('CC_CODE')+'"') ;
end ;

procedure TOT_TTVENTILTYPE.OnUpdateRecord ;
begin
  Inherited;
  if bCreate then begin
    if VH^.AnaCroisaxe then ParamVentilCroise('TY',GetField('CC_CODE'),taModif,True)
    else ParamVentil('TY',GetField('CC_CODE'),'12345',taModif,True);
    bCreate := False;
  end;
end;

procedure TOT_TTVENTILTYPE.OnAfterUpdateRecord ;
begin
  Inherited ;

end ;

procedure TOT_TTVENTILTYPE.OnClose ;
begin
  Inherited ;

end ;

procedure TOT_TTVENTILTYPE.OnArgument(S : String ) ;
begin
  OkComplement := True;
  bComplement.Hint := TraduireMemoire('Paramétrage des ventilations');
  bComplement.GlobalIndexImage := 'Z0133_S16G1';
  Inherited ;
end ;

procedure TOT_TTVENTILTYPE.OnComplement(Prefixe, Tipe, Code : string) ;
begin
  Inherited ;
  if VH^.AnaCroisaxe then ParamVentilCroise('TY',GetField('CC_CODE'),taModif,True)
    else ParamVentil('TY',GetField('CC_CODE'),'12345',taModif,True);
end ;

Initialization
  registerclasses ( [ TOT_TTVENTILTYPE ] );

end.
