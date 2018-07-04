unit UtofGcEtabliss_mul;

interface
uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
			emul,efiche, eFichList,  MaineAGL, eqrs1,
{$ELSE}
      dbTables,DBGrids, db,qrs1,mul,fiche,FichList, FE_main,
{$ENDIF}
      forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,{ mul,} HDimension,HDB,UTOM,
      {Fiche, FichList,}AglInit,UTOB,Dialogs,Menus, M3FP, EntGC, {FE_main,} grids,LookUp,
      AglInitGC, utilPGI, UtilGC;

type
     TOF_GCETABLISS_MUL = Class (TOF)
        procedure ErgoGCS3 ;
        procedure OnArgument(Arguments : String) ; override ;
        procedure OnLoad  ;override ;
     END ;

implementation

procedure TOF_GCETABLISS_MUL.ErgoGCS3 ;
{$IFDEF CCS3}
Var C,T : TControl ;
{$ENDIF}
BEGIN
{$IFDEF CCS3}
C:=TControl(Ecran.FindComponent('ET_DATELIBRE1')) ;
if C<>Nil then
   BEGIN
   T:=TControl(C.Parent) ;
   if T is TTabSheet then else T:=TControl(T.Parent) ;
   if T is TTabSheet then TTabSheet(T).TabVisible:=False ;
   END ;
C:=TControl(Ecran.FindComponent('ET_LIBREET1')) ;
if C<>Nil then
   BEGIN
   T:=TControl(C.Parent) ;
   if T is TTabSheet then else T:=TControl(T.Parent) ;
   if T is TTabSheet then TTabSheet(T).TabVisible:=False ;
   END ;
{$ENDIF}
END ;

procedure TOF_GCETABLISS_MUL.OnArgument(Arguments : String) ;
var Nbr : integer;
Begin
inherited ;
Nbr := 0;
if (GCMAJChampLibre (Tform (ecran), (Ecran is TFQRS1), 'COMBO', 'ET_LIBREET', 10, '') = 0) then SetControlVisible('PCOMPLEMENT', False) ;
if (GCMAJChampLibre (Tform (ecran), (Ecran is TFQRS1), 'EDIT', 'ET_VALLIBRE', 3, '_') = 0) then SetControlVisible('GRVALEUR', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (Tform (ecran), (Ecran is TFQRS1), 'EDIT', 'ET_DATELIBRE', 3, '_') = 0) then SetControlVisible('GRDATE', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (Tform (ecran), (Ecran is TFQRS1), 'BOOL', 'ET_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GRDECISION', False) else Nbr := Nbr + 1;
if (Nbr = 0) then SetControlVisible('PZLIBRE', False) ;
ErgoGCS3 ;
// pas de création sur le FO
if ctxFO in V_PGI.PGIContexte then SetControlVisible('Binsert', False) ;
End ;


procedure TOF_GCETABLISS_MUL.OnLoad  ;
var xx_where : string ;
begin
inherited ;
if not ((Ecran is TFMul) or (Ecran is TFQRS1)) then exit ;
xx_where:='' ;

// Gestion des checkBox : booléens libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'ET_BOOLLIBRE', 3, '');

// Gestion des dates libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'ET_DATELIBRE', 3, '_');

// Gestion des valeurs libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'ET_VALLIBRE', 3, '_');

SetControlText('XX_WHERE',xx_where) ;

end ;

Initialization
registerclasses([TOF_GCETABLISS_MUL]);

end.
