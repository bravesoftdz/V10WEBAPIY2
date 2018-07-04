unit UtofDispo_Cube;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, UTOF, mul,
  dbTables, StdCtrls, HCtrls, HDimension, AglInitGC, M3FP, FE_Main, hmsgbox,
  EntGC,HEnt1,UtilPGI,UtilGC,UtilArticle ;

Type
     TOF_DISPO_CUBE = Class (TOF)
        procedure OnArgument(Arguments : String) ; override ;
        procedure OnLoad ; override ;
     end ;

implementation

{ TOF_DISPO_CUBE }

procedure TOF_DISPO_CUBE.OnArgument(Arguments: String);
var iCol : integer ;
begin
Inherited ;
if (ctxMode in V_PGI.PGIContexte) then
   SetControlProperty('GQ_DEPOT','Plus','GDE_SURSITE="X"');

// Paramétrage des libellés des familles, stat. article et dimensions
for iCol:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+InttoStr(iCol),Ecran);
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
    begin
    for iCol:=4 to 8 do ChangeLibre2('TGA2_FAMILLENIV'+InttoStr(iCol),Ecran);
    for iCol:=1 to 2 do ChangeLibre2('TGA2_STATART'+InttoStr(iCol),Ecran);
    end ;

// Paramétrage des libellés des tables libres
GCMAJChampLibre (TForm (Ecran), False, 'MULTICOMBO', 'GA_LIBREART', 10, '');

// Mise en forme des libellés des dates, booléans libres et montants libres
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False);
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False);
if (GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False);
end;

procedure TOF_DISPO_CUBE.OnLoad;
var xx_where : string ;
begin
xx_where:='' ;

// Gestion des checkBox : booléens libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'GA_BOOLLIBRE', 3, '');

// Gestion des dates libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'GA_DATELIBRE', 3, '_');

// Gestion des montants libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'GA_VALLIBRE', 3, '_');

SetControlText('XX_WHERE',xx_where) ;
end;

Initialization
RegisterClasses([TOF_Dispo_Cube]) ;
end.
