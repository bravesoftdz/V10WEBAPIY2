unit UTofMBOEdtArt_Mode;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,UtilPGI,
      HCtrls,HEnt1,HMsgBox,UTOF,  HDimension,HDB,UTOM,
       AglInit,UTOB,Dialogs,Menus, M3FP, EntGC,grids,LookUp,
{$IFDEF EAGLCLIENT}
      eFiche,eQRS1,utileAGL,Maineagl,eFichList,
{$ELSE}
      QRS1,db,dbTables,DBGrids,FichList,Fiche,FE_main,
{$ENDIF}
      AglInitGC,UtilArticle,UtilGC;

Type
     TOF_MBOEDTART = Class (TOF)
        public
        procedure OnArgument(st:String) ; override ;
        procedure OnLoad ; override ;
     END ;

implementation

procedure TOF_MBOEDTART.OnArgument(st:String) ;
var Nbr, iCol : integer ;
begin
Inherited ;

// Paramétrage des libellés des familles, stat. article et dimensions
for iCol:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+InttoStr(iCol),Ecran);
ChangeLibre2('TGA_COLLECTION',Ecran);
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
    begin
    for iCol:=4 to 8 do ChangeLibre2('TGA2_FAMILLENIV'+InttoStr(iCol),Ecran);
    for iCol:=1 to 2 do ChangeLibre2('TGA2_STATART'+InttoStr(iCol),Ecran);
    end ;

// Paramétrage des libellés des tables libres
Nbr := 0;
if (GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False) ;

 if (ctxMode in V_PGI.PGIContexte) then
   begin
      // Mise en forme des libellés des dates, booléans libres et montants libres
      if (GCMAJChampLibre (TForm (Ecran), True, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else Nbr := Nbr + 1;
      if (GCMAJChampLibre (TForm (Ecran), True, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else Nbr := Nbr + 1;
      if (GCMAJChampLibre (TForm (Ecran), True, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else Nbr := Nbr + 1;
      if (Nbr = 0) then SetControlVisible('PZONESLIBRES', False) ;
   end ; // if (ctxMode in V_PGI.PGIContexte) then

end ;

procedure TOF_MBOEDTART.OnLoad  ;
var xx_where : string ;
begin
inherited ;
if not (Ecran is TFQRS1) then exit ;
xx_where:='' ;

// Gestion des checkBox : booléens libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'GA_BOOLLIBRE', 3, '');

// Gestion des dates libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'GA_DATELIBRE', 3, '_');

// Gestion des montants libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'GA_VALLIBRE', 3, '_');

SetControlText('XX_WHERE',xx_where) ;
end ;

Initialization
registerclasses([TOF_MBOEDTART]) ;
end.

