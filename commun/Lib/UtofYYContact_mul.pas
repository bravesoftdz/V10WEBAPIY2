unit UtofYYContact_mul;

interface
uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      eMul,
{$ELSE}
      {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}DBGrids, db,Fiche, FichList, FE_main,HDB,mul, 
{$ENDIF}
      forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,
      AglInit,UTOB,Dialogs,Menus, M3FP, EntGC, grids,LookUp,
      AglInitGC, utilPGI, UtilGC;

type 
     TOF_YYCONTACT_MUL = Class (TOF)
        procedure OnArgument(Arguments : String) ; override ;
        procedure OnLoad  ;override ;
     END ;

implementation

procedure TOF_YYCONTACT_MUL.OnArgument(Arguments : String) ;
var Nbr : integer;
Begin
inherited ;
Nbr := 0;
if (GCMAJChampLibre (Tform (ecran), False, 'COMBO', 'C_LIBRECONTACT', 10, '') = 0) then SetControlVisible('PCOMPLEMENT', False) ;
if (GCMAJChampLibre (Tform (ecran), False, 'EDIT', 'C_VALLIBRE', 3, '_') = 0) then SetControlVisible('GRVALEUR', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (Tform (ecran), False, 'EDIT', 'C_DATELIBRE', 3, '_') = 0) then SetControlVisible('GRDATE', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (Tform (ecran), False, 'BOOL', 'C_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GRDECISION', False) else Nbr := Nbr + 1;
if (Nbr = 0) then SetControlVisible('PZLIBRE', False) ;

End ;


procedure TOF_YYCONTACT_MUL.OnLoad  ;
var xx_where : string ;
begin
inherited ;
if not (Ecran is TFMul) then exit ;
xx_where:='' ;

// Gestion des checkBox : booléens libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'C_BOOLLIBRE', 3, '');

// Gestion des dates libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'C_DATELIBRE', 3, '_');

// Gestion des valeurs libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'C_VALLIBRE', 3, '_');

SetControlText('XX_WHERE',xx_where) ;

end ;

Initialization
registerclasses([TOF_YYCONTACT_MUL]);

end.
