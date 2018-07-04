unit UtofGenere_Actions;

interface
uses  Classes,forms,
      UTOF,UtilSelection,ParamSoc,
{$IFDEF EAGLCLIENT}
      MaineAGL;
{$ELSE}
      Fe_Main;
{$ENDIF}
Type
     TOF_Genere_Actions = Class (TOF)
     public
        procedure OnArgument(Arguments : String ) ; override ;
     END ;

Function RTLanceFiche_Genere_Actions(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation


Function RTLanceFiche_Genere_Actions(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Genere_Actions.OnArgument(Arguments : String ) ;
var F : TForm;
begin
inherited ;
F := TForm (Ecran);
if GetParamSocSecur('SO_RTGESTINFOS002',False) = True then
    MulCreerPagesCL(F,'NOMFIC=RTOPERATIONS');
end;

Initialization
registerclasses([TOF_Genere_Actions]);
end.
