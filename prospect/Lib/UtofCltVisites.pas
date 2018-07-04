unit UtofCltVisites;

interface

uses  Controls,Classes,sysutils,
      HCtrls,HEnt1,UTOF,Windows,
{$IFDEF EAGLCLIENT}
      eQRS1 ,MaineAGL
{$ELSE}

      QRS1,FE_Main
{$ENDIF}
        ;
Type
     Tof_CltVisites = Class (TOF)

     private

     public
       procedure OnArgument (Arguments : String ); override ;
       procedure OnLoad ; override ;
       procedure OnUpdate ; override ;
       procedure OnClose ; override ;
       procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
     END ;

Const
	// libellés des messages
	TexteMessage: array[1..1] of string 	= (
          {1}         'Vous devez sélectionner un type d''action.'
                 );

procedure ZoomListeTiers (CodeRep : string);

var TypeAction,NatureAuxi,DateActionDeb,DateActionFin : string;

implementation

procedure TOF_CltVisites.OnArgument (Arguments : String );
begin
inherited ;
Arguments := uppercase(Arguments);
if Arguments = 'COMMERCIAL' then
   begin
   TFQRS1(Ecran).CodeEtat:='RSR';
   TFQRS1(Ecran).NatureEtat:='RSR';
   SetControlVisible ('T_REPRESENTANT',True);
   SetControlVisible ('TT_REPRESENTANT',True);
   Ecran.Caption:= TraduireMemoire('Statistiques actions par commercial');
   updatecaption(Ecran);
   Ecran.HelpContext:=111000500 ;
   end
   else
   Ecran.HelpContext:=111000501 ;
if (ecran <> Nil)then TFQrs1(ecran).OnKeyDown:=FormKeyDown ;
end;

procedure TOF_CltVisites.OnLoad;
Var DateDeb : TDateTime;
begin
Inherited;
if GetControlText('TYPEACTION') = '' then
    BEGIN
    LastError:=1;
    LastErrorMsg:=TexteMessage[LastError] ;
    SetControlText('CODEACTION','');
    END;
DateDeb := StrToDate(GetControlText('DATEACTION'));
SetControlText('DATDEB',usdateTime(DateDeb));
DateDeb := StrToDate(GetControlText('DATEACTION_'));
SetControlText('DATFIN',usdateTime(DateDeb));
SetControlText('CODEACTION',GetControlText('TYPEACTION'));
// Chargement des critères de sélection en variables globales pour le zoom
TypeAction := GetControlText('TYPEACTION');
DateActionDeb := GetControlText('DATEACTION');
DateActionFin := GetControlText('DATEACTION_');
NatureAuxi := GetControlText('T_NATUREAUXI_');
end;

procedure TOF_CltVisites.OnUpdate;
begin
Inherited;
end;

procedure TOF_CltVisites.OnClose;
begin
TypeAction := '';
DateActionDeb := '';
DateActionFin := '';
NatureAuxi := '';
Inherited;
end;

procedure TOF_CltVisites.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  Case Key of
    VK_F10,VK_F9 :  if GetControlText ('TYPEACTION') = '' then key :=0;
  end;
if (ecran <> nil) then TFQrs1(ecran).FormKeyDown(Sender,Key,Shift);
end;

procedure ZoomListeTiers(CodeRep : string);
var St : string;
begin
St := 'T_REPRESENTANT='+CodeRep+';TYPEACTION='+TypeAction+';T_NATUREAUXI_='+NatureAuxi;
St := St + ';DATEACTION='+DateActionDeb+';DATEACTION_='+DateActionFin;
AGLLanceFiche('RT','RTPROSSACT_ZOOM',St,'','');
end;

Initialization
registerclasses([TOF_CltVisites]);


end. 
