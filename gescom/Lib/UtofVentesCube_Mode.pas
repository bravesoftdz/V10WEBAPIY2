unit UtofVentesCube_Mode;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, UTOF, mul,
  dbTables, StdCtrls, HCtrls, HDimension, AglInitGC, M3FP, FE_Main, hmsgbox,
  EntGC, HEnt1, utilPGI, UtilGC;

Type
     TOF_VENTESCUBE_MODE = Class (TOF)
        procedure OnArgument(Arguments : String) ; override ;
        procedure OnLoad ; override ;
     end ;

implementation

{ TOF_VENTESCUBE_MODE }

procedure TOF_VENTESCUBE_MODE.OnArgument(Arguments: String);
var FF : TForm;
begin
Inherited ;
FF:=TForm(Ecran) ;

if (ctxMode in V_PGI.PGIContexte) then
   SetControlProperty('GL_DEPOT','Plus','GDE_SURSITE="X"');

if Arguments = 'VEN' then
    begin
    //if (VH_GC.GCSeria = True or V_PGI.VersionDemo = True)
    //   then THValComboBox(FF.FindComponent('GL_NATUREPIECEG')).Value := 'FAC'
    //   else THValComboBox(FF.FindComponent('GL_NATUREPIECEG')).Value := 'FFO';
    if (VH_GC.GCSeria = True) or (V_PGI.VersionDemo = True)
       then SetControlText('GL_NATUREPIECEG','FAC;AVC;FFO')
       else SetControlText('GL_NATUREPIECEG','FFO');
    end else
    begin
    //THValComboBox(FF.FindComponent('GL_NATUREPIECEG')).Value := 'BLF';
    SetControlText('GL_NATUREPIECEG','BLF');
    end;

// Paramétrage des libellés des tables libres
GCMAJChampLibre (FF, False, 'COMBO', 'GL_LIBREART', 10, '');
end;

procedure TOF_VENTESCUBE_MODE.OnLoad;
begin
end;

Initialization
RegisterClasses([TOF_VentesCube_Mode]) ;
end.
