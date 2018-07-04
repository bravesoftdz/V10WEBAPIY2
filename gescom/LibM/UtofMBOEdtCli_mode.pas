unit UtofMBOEdtCli_mode;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,UtilPGI,
      HCtrls,HEnt1,HMsgBox,UTOF,  HDimension,HDB,UTOM,
       AglInit,UTOB,Dialogs,Menus, M3FP, EntGC,grids,LookUp,HTB97,
{$IFDEF EAGLCLIENT}
      eFiche,eQRS1,utileAGL,Maineagl,eFichList,
{$ELSE}
      QRS1,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids,FichList,Fiche,FE_main,
{$ENDIF}
      AglInitGC, UtilGC;

type
     TOF_MBOEDTCLI_MODE = Class (TOF)
        procedure OnArgument(Arguments : String) ; override ;
        procedure OnLoad  ;override ;
     private
				//uniquement en line
        //procedure CacheFonctionsNonLine;

     END ;

implementation

//uniquement en line
{*
procedure TOF_MBOEDTCLI_MODE.CacheFonctionsNonLine;
begin
  if TTabSheet(GetControl('PTABLESLIBRES')) <> nil then TTabSheet(GetControl('PTABLESLIBRES')).TabVisible := false;
  if TTabSheet(GetControl('PZONESLIBRES')) <> nil then TTabSheet(GetControl('PZONESLIBRES')).TabVisible := false;
  if THValComboBox(GetControl('_TYPECLI')) <> nil then THValComboBox(GetControl('_TYPECLI')).visible := false;
  if THlabel(GetControl('TTYPECLI')) <> nil then THlabel(GetControl('TTYPECLI')).visible := false;
  if THValComboBox(GetControl('T_TARIFTIERS')) <> nil then THValComboBox(GetControl('T_TARIFTIERS')).visible := false;
  if THlabel(GetControl('TT_TARIFTIERS')) <> nil then THlabel(GetControl('TT_TARIFTIERS')).visible := false;
  if TCheckBox(GetControl('T_FERME')) <> nil then
  begin
    TCheckBox(GetControl('T_FERME')).visible := false;
    TCheckBox(GetControl('T_FERME')).State := cbGrayed;
  end;
  if THlabel(GetControl('TEtat')) <> nil then THlabel(GetControl('TEtat')).visible := false;
  if THValComboBox(GetControl('Fetat')) <> nil then THValComboBox(GetControl('Fetat')).visible := false;
  if TTabSheet(GetControl('OPTION')) <> nil then TTabSheet(GetControl('OPTION')).Tabvisible := false;
  SetControlText ('_TYPECLI','');
  TFQRS1(Ecran).CodeEtat := 'BCL';
end;
*}

procedure TOF_MBOEDTCLI_MODE.OnArgument(Arguments : String) ;
var Nbr : integer;
Begin
inherited ;
{$IFDEF NOMADE}
SetControlProperty ( 'GL_NATUREPIECEG', 'PLUS', 'AND (GPP_NATUREPIECEG IN ("CC","DE") OR GPP_TRSFVENTE="X")' ) ;
{$ENDIF}
Nbr := 0;
if (GCMAJChampLibre (Tform (ecran), True, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False) ;
if (GCMAJChampLibre (Tform (ecran), True, 'EDIT', 'YTC_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (Tform (ecran), True, 'EDIT', 'YTC_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else Nbr := Nbr + 1;
if (GCMAJChampLibre (Tform (ecran), True, 'BOOL', 'YTC_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else Nbr := Nbr + 1;
if (Nbr = 0) then SetControlVisible('PZONESLIBRES', False) ;
//uniquement en line
//	CacheFonctionsNonLine;

End ;


procedure TOF_MBOEDTCLI_MODE.OnLoad  ;
var xx_where : string ;
begin
  inherited ;
  if not (Ecran is TFQRS1) then exit ;
  if Ecran.Name = 'GCEDTHITCLI_MODE' then
    xx_where:='T_NATUREAUXI="CLI" and VTR_NATUREPIECEG="FFO"' else
  if Ecran.Name = 'EDTCPTCLI_MODE' then
    xx_where:='T_NATUREAUXI="CLI" and GPE_NATUREPIECEG="FFO" and MP_TYPEMODEPAIE<>"007"' else
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
registerclasses([TOF_MBOEDTCLI_MODE]);

end.
