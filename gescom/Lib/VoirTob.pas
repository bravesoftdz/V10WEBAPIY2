unit VoirTob;

interface

uses  Classes,HCtrls,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      FE_Main,
{$ENDIF}
      AGLInit, UTOF, Utob,Vierge;

Type

     TOF_VOIRTOB = Class (TOF)
     public
        procedure OnLoad ; override ;
     END ;

procedure GCVoirTob ( ToTob : Tob );

implementation

procedure GCVoirTob ( ToTob : Tob );
begin
TheTob:=Totob ;
AGLLanceFiche('GC','GCVOIRTOB','','','');
end ;

Procedure TOF_VOIRTOB.Onload ;
begin
inherited ;
LaTob.PutGridDetail(THGRID(GetControl('GS')),True,True,'',True);
end;
Initialization
registerclasses([TOF_VOIRTOB]);
end.
 