unit UtofRTOutlook;

interface

uses Classes,Controls,UTOF,OutlookView_TLB
     ,M3FP,forms,vierge
     ;

Type
    TOF_RTOutlook = Class (TOF)
    Private
      OVCtl1: TOVCtl;
    Public
      procedure GoToday ;
      procedure OnClose ; override ;
      procedure OnArgument(Arguments : String ) ; override ;
    END;

implementation

procedure TOF_RTOutlook.OnArgument(Arguments : String ) ;
begin
inherited;
OVCtl1:= TOVCtl.create(Ecran);
OVCtl1.parent:=TWinControl(getcontrol('AGENDA'));
OVCtl1.folder:='Calendrier';
OVCtl1.align:=alClient;
end;

procedure TOF_RTOutlook.GoToday ;
begin
OVCtl1.GoToToday;
end;


procedure TOF_RTOutlook.OnClose ;
begin
OVCtl1.free;
inherited;
end;

Procedure AGLRTOLKGoToday( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_RTOutlook(totof).GoToday;
end;

Initialization
registerclasses([TOF_RTOutlook]);
RegisterAglProc( 'RTOLKGoToday', TRUE , 0, AGLRTOLKGoToday);
end.
