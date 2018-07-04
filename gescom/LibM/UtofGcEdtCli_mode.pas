unit UtofGcEdtCli_mode;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,UtilPGI,
      HCtrls,HEnt1,HMsgBox,UTOF,  HDimension,HDB,UTOM,
       AglInit,UTOB,Dialogs,Menus, M3FP, EntGC,grids,LookUp,
{$IFDEF EAGLCLIENT}
      eFiche,eQRS1,utileAGL,Maineagl,eFichList,
{$ELSE}
      QRS1,db,dbTables,DBGrids,FichList,Fiche,FE_main,
{$ENDIF}
      AglInitGC, UtilGC;

type 
     TOF_GCEDTCLI_MODE = Class (TOF)
        procedure OnArgument(Arguments : String) ; override ;
        procedure OnLoad  ;override ;
     END ;

implementation

procedure TOF_GCEDTCLI_MODE.OnArgument(Arguments : String) ;
var Nbr : integer;
Begin
inherited ;
Nbr := 0;
if (GCMAJChampLibre (Tform (ecran), True, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False) ;
if (GCMAJChampLibre (Tform (ecran), True, 'EDIT', 'YTC_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (Tform (ecran), True, 'EDIT', 'YTC_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (Tform (ecran), True, 'BOOL', 'YTC_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else Nbr := Nbr + 1;
if (Nbr = 0) then SetControlVisible('PZONESLIBRES', False) ;
End ;


procedure TOF_GCEDTCLI_MODE.OnLoad  ;
var xx_where : string ;
begin
inherited ;
if not (Ecran is TFQRS1) then exit ;
if Ecran.Name = 'GCEDTHITCLI_MODE' then
   xx_where:='T_NATUREAUXI="CLI" and VTR_NATUREPIECEG="FFO"'
else
   xx_where:='T_NATUREAUXI="CLI"';

// Gestion des checkBox : booléens libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'YTC_BOOLLIBRE', 3, '');

// Gestion des dates libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'YTC_DATELIBRE', 3, '_');

// Gestion des valeurs libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'YTC_VALLIBRE', 3, '_');

SetControlText('XX_WHERE',xx_where) ;

end ;

Initialization
registerclasses([TOF_GCEDTCLI_MODE]);

end.
 