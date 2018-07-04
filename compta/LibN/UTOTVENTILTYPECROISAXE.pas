{***********UNITE*************************************************
Auteur  ...... : Guillon Stéphane
Créé le ...... : 28/01/05
Modifié le ... :   /  /
Description .. : Source TOT de la TABLE : TTVENTILTYPECROISAXE
Mots clefs ... : TOT;TTVENTILTYPE
*****************************************************************}
Unit UTOTVENTILTYPECROISAXE;

Interface

uses
  Classes, // TMemoryStream, HexToBin
  UTot,
  HPanel,  // THPanel
  HCtrls,  // ExecuteSQL,
  HEnt1,   // taModif
  HTB97,   // TToolbarButton97
  Ent1, //VH^
  CPVENTILTYPECROIS_TOF   // ParamVentilCroise
  ;

Type
  TOT_TTVENTILTYPECROISAXE = Class ( TOT )
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

procedure TOT_TTVENTILTYPECROISAXE.OnNewRecord ;
begin
  Inherited ;
  bCreate := True;
end ;

procedure TOT_TTVENTILTYPECROISAXE.OnDeleteRecord ;
begin
  Inherited ;
  ExecuteSql('DELETE FROM VENTIL WHERE V_NATURE LIKE "TY%" AND V_COMPTE="'+GetField('CC_CODE')+'"') ;
end ;

procedure TOT_TTVENTILTYPECROISAXE.OnUpdateRecord ;
begin
  Inherited;
  if bCreate then begin
    ParamVentilCroise('TY',GetField('CC_CODE'),taModif,True);
    bCreate := False;
  end;
end;

procedure TOT_TTVENTILTYPECROISAXE.OnAfterUpdateRecord ;
begin
  Inherited ;

end ;

procedure TOT_TTVENTILTYPECROISAXE.OnClose ;
begin
  Inherited ;

end ;

procedure ChargeGlyph(bb : TToolbarButton97);
const
  Data = '424D6601000000000000760000002800000014000000140000000100'+
         '040000000000F000000000000000000000001000000010000000000000000000'+
         '8000008000000080800080000000800080008080000080808000C0C0C0000000'+
         'FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00888888888888'+
         '8888888800008888888888888888888800008888888888888888888800008888'+
         '8884444488888888000088888884FFF488888888000088888804444488888888'+
         '00008888808888888888888800004444488444448844444800004FFF4004BBB4'+
         '004FFF4800004444488444448844444800008888808888888888888800008888'+
         '8804444488444448000088888884FFF4004FFF48000088888884444488444448'+
         '0000888888888888088888880000888888888888804444480000888888888888'+
         '884FFF4800008888888888888844444800008888888888888888888800008888'+
         '88888888888888880000';
var
  s   : TMemoryStream  ;
begin
  s := TMemoryStream .Create;
  try
    S.SetSize(Length(Data) div 2);
    HexToBin(PAnsiChar(Data),s.Memory,s.Size);
    s.Seek(0,0);
    bb.Glyph.LoadFromStream(s);
  finally
    s.Free;
  end;
end;

procedure TOT_TTVENTILTYPECROISAXE.OnArgument(S : String ) ;
begin
  OkComplement := True;
  bComplement.Hint := TraduireMemoire('Paramétrage des ventilations');
  bComplement.NumGlyphs := 1;
  ChargeGlyph(bComplement);
  Inherited ;
end ;

procedure TOT_TTVENTILTYPECROISAXE.OnComplement(Prefixe, Tipe, Code : string) ;
begin
  Inherited ;
  ParamVentilCroise('TY',Code,taModif,True) ;
end ;

Initialization
  registerclasses ( [ TOT_TTVENTILTYPECROISAXE ] );

end.
