unit UtofActions_TV;

interface
uses  Classes,forms,sysutils,
      HEnt1,UTOF,M3FP, UTobView, TVProp
{$IFDEF EAGLCLIENT}
      ,MainEAGL, emul
{$ELSE}
      ,Fe_Main,mul
{$ENDIF}
      ,UtilSelection,UtilRT;

Function RTLanceFiche_Actions_TV(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
     TOF_Actions_TV = Class (TOF)
     private
         TobViewer1: TTobViewer;
         procedure ActTVOnDblClickCell(Sender: TObject ) ;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
     END ;



implementation

Function RTLanceFiche_Actions_TV(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;


procedure TOF_Actions_TV.OnArgument(Arguments : String ) ;
var F : TForm;
    NomForme : String;
begin
inherited ;
F := TForm (Ecran);
MulCreerPagesCL(F);
NomForme:=F.Name;

F.Name:='RTACTIONS';
MulCreerPagesCL(F);
F.Name:=NomForme;

TobViewer1:=TTobViewer(getcontrol('TV'));
TobViewer1.OnDblClick:= ActTVOnDblClickCell ;
end;

procedure TOF_Actions_TV.OnLoad;
var xx_where : string ;
begin
inherited;
SetControlText('XX_WHERE',RTXXWhereConfident('CON')) ;
end;

procedure TOF_Actions_TV.ActTVOnDblClickCell(Sender: TObject );
var Stchaine : string;
begin
with TTobViewer(sender) do
    begin
    if (ColName[CurrentCol] = 'RAC_AUXILIAIRE') or (ColName[CurrentCol] = 'T_LIBELLE') or (ColName[CurrentCol] = 'RAC_TIERS')then
      V_PGI.DispatchTT (28,taConsult ,AsString[ColIndex('RAC_AUXILIAIRE'), CurrentRow], '','')
    else if (ColName[CurrentCol] = 'RAC_INTERVENANT') or (ColName[CurrentCol] = 'ARS_LIBELLE') then
//      V_PGI.DispatchTT (9,taConsult ,AsString[ColIndex('RAC_INTERVENANT'), CurrentRow], '','')
      V_PGI.DispatchTT (6,taConsult ,AsString[ColIndex('RAC_INTERVENANT'), CurrentRow], '','')
    else if (ColName[CurrentCol] = 'RAC_PROJET') then
        begin
        StChaine := AsString[ColIndex('RAC_PROJET'), CurrentRow];
        if Stchaine <> '' then
           V_PGI.DispatchTT (30,taConsult ,AsString[ColIndex('RAC_PROJET'), CurrentRow], '','')
        end
    else if (ColName[CurrentCol] = 'RAC_OPERATION') then
        begin
        StChaine := AsString[ColIndex('RAC_OPERATION'), CurrentRow];
        if Stchaine <> '' then
           V_PGI.DispatchTT (23,taConsult ,AsString[ColIndex('RAC_OPERATION'), CurrentRow], '','')
        end
    else
        begin
        if copy(ColName[CurrentCol],1,3) = 'RCH' then
           V_PGI.DispatchTT (24,taConsult ,IntToStr(AsInteger[ColIndex('RCH_NUMERO'), CurrentRow]), '','')
        else
           V_PGI.DispatchTT (22,taConsult ,AsString[ColIndex('RAC_AUXILIAIRE'), CurrentRow]+';'+IntToStr(AsInteger[ColIndex('RAC_NUMACTION'), CurrentRow]), '','');
        end;
    end;
end;

Initialization
registerclasses([TOF_Actions_TV]);

end.
