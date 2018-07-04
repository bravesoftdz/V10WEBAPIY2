unit MDispGAUtil;

interface
uses  SysUtils,Ent1,HStatus,HEnt1,
{$IFDEF EAGLCLIENT}
      UtileAGL ,
{$ELSE}
      MajTable ,
{$ENDIF}
   HCtrls;

Procedure TripoteStatusAF ;


implementation


Procedure TripoteStatusAF ;
Var St,AJ : String ;
BEGIN
AJ:=VH^.CPStatusBarre ;
if ((AJ='') or (AJ='AUC')) then Exit ;
St:=GetDefStatus ;
if AJ='PIV' then St:=St+'   '+RechDom('TTDEVISETOUTES',V_PGI.DevisePivot,False)+' ('+V_PGI.DevisePivot+')' else
if AJ='CON' then St:=St+'   '+DateToStr(V_PGI.DateEntree) else
if AJ='EXO' then St:=St+'   '+RechDom('TTEXERCICE',VH^.Encours.Code,False)+' ('+VH^.Encours.Code+')' else
if AJ='CLO' then St:=St+'   '+DateToStr(VH^.DateCloturePer) else
if AJ='BAS' then St:=St+'   '+GetDBPathName(TRUE) else
if AJ='VER' then St:=St+'   '+V_PGI.NumVersion+' du '+DateToStr(V_PGI.DateVersion) else ;
ChgStatus(St) ;
END ;


end.
