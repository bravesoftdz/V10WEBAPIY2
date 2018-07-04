unit UtofGcFoumul_mode;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,utilPGI,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,HDB,UTOM,
      AglInit,UTOB,Dialogs,Menus, M3FP, EntGC, grids,LookUp,
{$IFDEF EAGLCLIENT}
      eFichList,eFiche,emul,MaineAgl,eQRS1,
{$ELSE}
      QRS1,FichList,Fiche,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,DBGrids, mul,Fe_Main,
{$ENDIF}
      AglInitGC, UtilGC;

type
     TOF_GCFOUMUL_MODE = Class (TOF)
        procedure OnArgument(Arguments : String) ; override ;
        procedure OnLoad  ;override ;
     END ;

implementation

procedure TOF_GCFOUMUL_MODE.OnArgument(Arguments : String) ;
var Nbr,x : integer;
    stArg,critere,ChampMul : string;
Begin
inherited ;
Nbr:=0;
if (GCMAJChampLibre (Tform (ecran), False, 'COMBO', 'YTC_TABLELIBREFOU', 3, '') = 0) then SetControlVisible('GB_TABLE', False) else inc(Nbr);
if (GCMAJChampLibre (Tform (ecran), False, 'EDIT', 'YTC_VALLIBREFOU', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(Nbr);
if (GCMAJChampLibre (Tform (ecran), False, 'EDIT', 'YTC_DATELIBREFOU', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(Nbr);
if (Nbr=0) then SetControlVisible('PZONESLIBRES', False) ;

// initialisation de la variable Where
THEdit(GetControl('XX_WHERE')).Text:='T_NATUREAUXI="FOU"' ;

// Initialisation car StringToAction('') retourne par défaut taConsult.
if GetControl('TYPEACTION') <> nil then
  SetControlText('TYPEACTION','ACTION=MODIFICATION');

stArg:=Arguments;
Repeat
    Critere:=UpperCase(Trim(ReadTokenSt(stArg)));
    if Critere<>'' then
        BEGIN
        x:=pos('=',Critere);
        if x<>0 then
           BEGIN
           ChampMul:=copy(Critere,1,x-1);
           //ValMul:=copy(Critere,x+1,length(Critere));
           END;
           if ChampMul='ACTION' then SetControlText('TYPEACTION',Critere);
        END;
until Critere='';

if GetControl('bInsert')<>nil then
    SetControlVisible('bInsert',StringToAction(GetControlText('TYPEACTION'))<>taConsult);

if not (Ecran is TFQRS1) then exit ;
Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
if Critere='ETIQ' then
   begin
        TFQRS1(Ecran).CodeEtat:='GEF';
        TFQRS1(Ecran).NatureEtat:='GEF';
   end;

//uniquement en line
{*if getcontrol('option') <> nil then
   begin
   SetControlVisible('PZONESLIBRES', False);
   SetControlVisible('Avances', False);
   SetControlVisible('option', False);
   end;
*}

End ;


procedure TOF_GCFOUMUL_MODE.OnLoad  ;
var xx_where : string ;
begin
inherited ;

if not ((Ecran is TFMul) or (Ecran is TFQRS1)) then exit ;
xx_where:='' ;

// Gestion des dates libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'YTC_DATELIBREFOU', 3, '_');

// Gestion des tables libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'COMBO', 'YTC_TABLELIBREFOU', 3, '');

// Gestion des valeurs libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'YTC_VALLIBREFOU', 3, '_');

SetControlText('XX_WHERE',xx_where) ;

end ;

Initialization
registerclasses([TOF_GCFOUMUL_MODE]);

end.
