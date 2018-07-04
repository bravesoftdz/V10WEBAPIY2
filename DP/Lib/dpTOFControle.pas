Unit dpTOFControle ;
// TOF de la fiche DP CONTROLE_FISCAL


Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF, HDB,
{$IFDEF EAGLCLIENT}
     eMul, MaineAGL,
{$ELSE}
     Mul, FE_Main,
{$ENDIF}
     HTB97, AGLInit, HQry;

Type
  TOF_DPCONTROLE = Class (TOF)
    procedure OnArgument (stArgument : String ) ; override ;
  private
    TypeCtrl : String; // FIS ou SOC
    procedure FListe_OnDblClick(Sender: TObject);
    procedure BINSERT_OnClick(Sender: TObject);
    procedure BCONSULT_OnClick(Sender: TObject);
    procedure BDELETE_OnClick(Sender: TObject);
  end ;


/////////////// IMPLEMENTATION ////////////////
Implementation

uses dpOutils,

     {$IFDEF VER150}
     Variants,
     {$ENDIF}

     DpJurOutilsEve, UDossierSelect;


procedure TOF_DPCONTROLE.OnArgument (stArgument : String ) ;
begin
  Inherited ;
  TypeCtrl := ReadTokenSt (stArgument);
  if TypeCtrl='' then TypeCtrl := 'FIS';

  THValComboBox(GetControl('DCL_DETAILCTRL')).Plus := ' AND CO_ABREGE = "'+TypeCtrl+'"';
  // en eagl, casté en thgrid par HDB
  THDBGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;
  TToolbarButton97(GetControl('BINSERT')).OnClick := BINSERT_OnClick;
  TToolbarButton97(GetControl('BCONSULT')).OnClick := BCONSULT_OnClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := BDELETE_OnClick;
end ;


procedure TOF_DPCONTROLE.FListe_OnDblClick(Sender: TObject);
var St : String;
begin
  if VarIsNull(GetField('DCL_GUIDPER')) then exit;
  St := GetField('DCL_GUIDPER')+';'+IntToStr(GetField('DCL_NOORDRE'));
  AGLLanceFiche('DP','DETCTRL_FISCAL', St, St,'ACTION=MODIFICATION;'+TypeCtrl) ;
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


procedure TOF_DPCONTROLE.BINSERT_OnClick(Sender: TObject);
begin
  AGLLanceFiche('DP', 'DETCTRL_FISCAL','','', 'ACTION=CREATION;'+TypeCtrl) ;
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


procedure TOF_DPCONTROLE.BCONSULT_OnClick(Sender: TObject);
begin
  FListe_OnDblClick(Sender);
end;


procedure TOF_DPCONTROLE.BDELETE_OnClick(Sender: TObject);
begin
  SupprimeListeEnreg(TFMul(Ecran), 'DPCONTROLE');
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
  // si on a supprimé le dernier

  if Not ExisteSQL('SELECT 1 FROM DPCONTROLE WHERE DCL_GUIDPER="'+VH_Doss.GuidPer+'"') then
   ExecuteSql ('UPDATE DPFISCAL SET DFI_CTRLFISC="-" WHERE DFI_GUIDPER="'+VH_Doss.GuidPer+'"');
end;

Initialization
  registerclasses ( [ TOF_DPCONTROLE ] ) ;
end.


