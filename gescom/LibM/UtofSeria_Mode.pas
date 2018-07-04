unit UtofSeria_Mode;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, UTOF, mul,
  dbTables, HCtrls, HDimension, AglInitGC, M3FP, FE_Main, UtilArticle,hmsgbox, EntGC,
  HEnt1, UtilDispGC;

Type
     TOF_SERIA_MODE = Class (TOF)
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
     end ;

implementation

Uses utilGC;            

procedure TOF_SERIA_MODE.OnLoad;
begin
if (ctxMode in V_PGI.PGIContexte) then
   begin
    if VH_GC.BOSeria=True
       then SetControlText('BOSERIA','X')
       else SetControlEnabled('BOSERIA', False);
    if VH_GC.GCSeria=True
       then SetControlText('GCSERIA','X')
       else SetControlEnabled('GCSERIA', False);

    if (V_PGI.LaSerie=S3) then
       begin
       if VH_GC.NbEtablisSeria>=2
          then SetControlText('L1SERIA','X')
          else SetControlEnabled('L1SERIA', False);
       if VH_GC.NbEtablisSeria>=5
          then SetControlText('L2SERIA','X')
          else SetControlEnabled('L2SERIA', False);
       if VH_GC.NbEtablisSeria>=10
          then SetControlText('L3SERIA','X')
          else SetControlEnabled('L3SERIA', False);
       SetControlVisible('L4SERIA', False);
       end else
       begin
       if VH_GC.NbEtablisSeria>=10
          then SetControlText('L1SERIA','X')
          else SetControlEnabled('L1SERIA', False);
       if VH_GC.NbEtablisSeria>=20
          then SetControlText('L2SERIA','X')
          else SetControlEnabled('L2SERIA', False);
       if VH_GC.NbEtablisSeria>=50
          then SetControlText('L3SERIA','X')
          else SetControlEnabled('L3SERIA', False);
       if VH_GC.NbEtablisSeria>50
          then SetControlText('L4SERIA','X')
          else SetControlEnabled('L4SERIA', False);
       end;
   end;
end;

procedure TOF_SERIA_MODE.OnUpdate;
begin
if (ctxMode in V_PGI.PGIContexte) then
   begin
    if GetControlText('BOSERIA')='-'
       then VH_GC.BOSeria:=False
       else VH_GC.BOSeria:=True;
    if GetControlText('GCSERIA')='-'
       then VH_GC.GCSeria:=False
       else VH_GC.GCSeria:=True;
    if (VH_GC.BOSeria=False) and (VH_GC.GCSeria=False)
       then V_PGI.VersionDemo:= True
       else V_PGI.VersionDemo:= False;

    VH_GC.NbEtablisSeria:=0;
    if (V_PGI.LaSerie=S3) then
       begin
        if GetControlText('L1SERIA')<>'X' then VH_GC.NbEtablisSeria := 0
        else if GetControlText('L2SERIA')<>'X' then VH_GC.NbEtablisSeria := 2
        else if GetControlText('L3SERIA')<>'X' then VH_GC.NbEtablisSeria := 5
        else VH_GC.NbEtablisSeria := 10;
       end else
       begin
        if GetControlText('L1SERIA')<>'X' then VH_GC.NbEtablisSeria := 0
        else if GetControlText('L2SERIA')<>'X' then VH_GC.NbEtablisSeria := 10
        else if GetControlText('L3SERIA')<>'X' then VH_GC.NbEtablisSeria := 20
        else if GetControlText('L4SERIA')<>'X' then VH_GC.NbEtablisSeria := 50
        else VH_GC.NbEtablisSeria := 9999;
       end;
    if VH_GC.NbEtablisSeria = 0 then
       begin                   // Si GC uniquement 1 btq ne demande pas de séria btq
       if VH_GC.BOSeria=True then V_PGI.VersionDemo:= True
       else if CompteNbEtablissMode > 1 then V_PGI.VersionDemo:= True;
       end;

    ChargeModules_Mode ( False );
   end;
end;

Initialization
RegisterClasses([TOF_Seria_Mode]) ;
end.
