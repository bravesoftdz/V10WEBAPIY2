unit dpTOMControle;
// TOM de la table DPCONTROLE

interface
uses  StdCtrls,Controls,Classes,db,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,HDB,UTOM,UTOB,HTB97,
{$IFDEF EAGLCLIENT}
      eFiche,
{$ELSE}
      Fiche,
      DBCtrls,
      dbTables,
{$ENDIF}
      UDossierSelect,
      Spin,Dialogs;

Type
  TOM_DPCONTROLE = Class (TOM)
    procedure OnNewRecord  ; override ;
    procedure OnAfterDeleteRecord  ; override ;
    procedure OnLoadRecord ; override ;
    procedure OnUpdateRecord  ; override ;
    procedure OnChangeField(F: TField);  override;
    procedure OnArgument (stArgument : String ) ; override ;
  private
    DateCreation     : TDateTime;
    TypeCtrl         : String;
  end ;


////////////// IMPLEMENTATION //////////////////
implementation

uses dpOutils, dpTOFEvoluCapital, galOutil;


{ -------------- CONTROLES FISCAUX ------------ }
procedure TOM_DPCONTROLE.OnLoadRecord;
begin
  SetControlText('DCL_GUIDPER', VH_Doss.GuidPer);
  SetField('DCL_GUIDPER', VH_Doss.GuidPer);
  DateCreation := GetDateCreation (VH_Doss.GuidPer);
end;


procedure TOM_DPCONTROLE.OnNewRecord;
begin
  SetField('DCL_GUIDPER',VH_Doss.GuidPer);
  SetField('DCL_NOORDRE',IncrementeSeqNoOrdre('DCL','DPCONTROLE',VH_Doss.GuidPer)) ;
  SetField('DCL_TYPECTRL', TypeCtrl);
end;


procedure TOM_DPCONTROLE.OnAfterDeleteRecord  ;
begin
  // si on a supprimé le dernier
  if Not ExisteSQL('SELECT 1 FROM DPCONTROLE WHERE DCL_GUIDPER="'+VH_Doss.GuidPer+'"') then
    ExecuteSql ('UPDATE DPFISCAL SET DFI_CTRLFISC="-" WHERE DFI_GUIDPER="'+VH_Doss.GuidPer+'"');
end;


procedure TOM_DPCONTROLE.OnChangeField(F: TField);
var TX  : double;
begin
  if (F.FieldName='DCL_REDRENVISAG') AND (DS.State in [dsInsert,dsEdit]) then
    begin
    Tx := GetField ('DCL_REDRENVISAG');
    if (GetField ('DCL_REDRACCEPT') = 0.00) AND (Tx <> 0.00) then SetField ('DCL_REDRACCEPT', Tx);
    end;
end;


procedure TOM_DPCONTROLE.OnArgument(stArgument: String);
var action : string;
begin
  Inherited ;
  action := ReadTokenSt(stArgument);  // ACTION = ...
  TypeCtrl := ReadTokenSt(stArgument);
  if (TypeCtrl = '') then TypeCtrl := 'FIS';
  THValComboBox(GetControl('DCL_DETAILCTRL')).Plus := ' AND CO_ABREGE = "'+TypeCtrl+'"';
end;


procedure TOM_DPCONTROLE.OnUpdateRecord;
// contrôles de cohérence
var
  wA, wM, wJ : word;
begin
  if StrToDate(GetField('DCL_DATENOTIF')) < DateCreation then
    begin
    ErreurChamp('DCL_DATENOTIF', 'PGeneral', Self, 'Date de notification ne peut être inférieure à la date de création de l''enreprise') ;
    exit;
    end;
  if GetField('DCL_EXERCFIN') < GetField('DCL_EXERCDEB') then
    begin
    ErreurChamp('DCL_EXERCFIN', 'PGeneral', Self, 'La borne "De" ne peut être inférieure à la borne "à"') ;
    exit;
    end;
  DecodeDate(strtodate(GetField('DCL_DATENOTIF')), wA , wM, wJ);
  if GetField('DCL_EXERCFIN') > wA then
    begin
    ErreurChamp('DCL_EXERCFIN', 'PGeneral', Self, 'La borne "à" ne peut être supérieure à l''année de notification') ;
    exit;
    end;


  if (StrToDate(GetField('DCL_DATEFINREDR')) <> iDate1900) and
     (StrToDate(GetField('DCL_DATEFINREDR')) < StrToDate(GetField('DCL_DATENOTIF'))) then
    begin
    ErreurChamp('DCL_DATEFINREDR', 'PGeneral', Self, 'Date de fin de procédure ne peut être inférieure à la date de notification') ;
    exit;
    end;
  if (StrToDate(GetField('DCL_DATEPAIEMENT')) <> iDate1900) and
     (StrToDate(GetField('DCL_DATEPAIEMENT')) < StrToDate(GetField('DCL_DATENOTIF'))) then
    begin
    ErreurChamp('DCL_DATEPAIEMENT', 'PGeneral', Self, 'Date de paiement ne peut être inférieure à la date de notification');
    exit;
    end;

  //--- Force DFI_CTRLFISC à True (existence d'au moins un contrôle fiscal validé)
  ExecuteSql ('UPDATE DPFISCAL SET DFI_CTRLFISC="X" WHERE DFI_GUIDPER="'+GetField('DCL_GUIDPER')+'"');
end;


Initialization
registerclasses([TOM_DPCONTROLE]) ;
end.
