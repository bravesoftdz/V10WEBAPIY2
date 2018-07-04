unit UtofDispoArt_Mode;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,     //  Menus,UTOB,AglInit,ComCtrls,HDB,UTOM
      HCtrls,HEnt1,HMsgBox,UTOF,HDimension,
      Dialogs,M3FP,EntGC,
{$IFDEF EAGLCLIENT}
      emul, MaineAGL,
{$ELSE}
      Fiche, mul, dbTables,Fe_Main,                 //  DBGrids,db,FichList,
{$ENDIF}
      LookUp,AglInitGC,UtilGC,UtilArticle ;                             //   grids,

// Windows, Messages, , , Graphics,, utilPGI;

Type
     TOF_DISPOART_MODE = Class (TOF)
        procedure OnLoad ; override ;
        procedure OnArgument(Arguments : String) ; override ;
        procedure OnUpdate ; override ;
     end ;

implementation

procedure TOF_DISPOART_MODE.OnLoad;
var xx_where : string ;
begin
if (ctxMode in V_PGI.PGIContexte) then
   SetControlProperty('GQ_DEPOT','Plus','GDE_SURSITE="X"');

xx_where:='' ;

// Gestion des checkBox : booléens libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'GA_BOOLLIBRE', 3, '');

// Gestion des dates libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'GA_DATELIBRE', 3, '_');

// Gestion des montants libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'GA_VALLIBRE', 3, '_');

SetControlText('XX_WHERE',xx_where) ;
end;

procedure TOF_DISPOART_MODE.OnArgument(Arguments: String);
var Nbr : integer ;
begin
Inherited ;
// Paramétrage des libellés des tables libres
Nbr := 0;
if (GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False) ;

// Mise en forme des libellés des dates, booléans libres et montants libres
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else Nbr := Nbr + 1;
if (Nbr = 0) then SetControlVisible('PZONESLIBRES', False) ;
end;


procedure TOF_DISPOART_MODE.OnUpdate;
//var   i : integer;
//      CC : THLabel;
//      FF : TFMUL;
begin
{
FF:=TFMul(Ecran) ;

for i:=0 to FF.FListe.Columns.Count-1 do
    begin
    // Reprise du libellé Famille niv 1, niv 2, niv 3 dans le bandeau entête de colonne
    if copy(FF.FListe.Columns[i].Title.caption,1,14)='Famille niveau' then
        begin
        CC:=THLabel(FF.FindComponent('TGA_FAMILLENIV'+copy(FF.FListe.Columns[i].Title.caption,16,1))) ;
        FF.Fliste.columns[i].Field.DisplayLabel:=CC.Caption ;
        end;
    // Personalisation du libellé de la colonne Dépôt en fonction du paramétrage mono-dépôt ou multi-dépôts
    if (not VH_GC.GCMultiDepots) and (copy(FF.FListe.Columns[i].Title.caption,1,5)='Dépôt') then
        begin
        FF.Fliste.columns[i].Field.DisplayLabel:='Etablis.' ;
        end;
    end;
}
end;

// Parms[1] = GetChamp('GQ_ARTICLE') -> réf. article générique
// Parms[2] = STOCK=X;ARTGEN=...;DEPOT=...;CLOTURE=...
Procedure TOFAfficheDispoGenerique (Parms : Array of variant; nb: integer);
var st1, st2 : string;
begin
st1 := Parms [1];
st2 := CodeArticleUnique2(st1, '');
if (Parms[3]='GEN') or (Parms[3]='DIM')
then AGLLanceFiche('GC','GCARTICLE', '', st2, 'ACTION=CONSULTATION;'+Parms[2]) ;
end;


Procedure InitTOFDispoArt_Mode ;
begin
RegisterAglProc('AfficheDispoGenerique', True , 1, TOFAfficheDispoGenerique) ;
End ;

Initialization
RegisterClasses([TOF_DispoArt_Mode]) ;
InitTOFDispoArt_Mode ;
end.
