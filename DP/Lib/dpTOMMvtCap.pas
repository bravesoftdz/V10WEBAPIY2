unit dpTOMMvtCap;
// TOM de la table DPMVTCAP (fiche DP MOUVCAPITAL)

interface
uses  StdCtrls,Controls,Classes,db,forms,sysutils,dbTables,ComCtrls,
      HCtrls,HEnt1,HMsgBox,HDB,UTOM,UTOB,HTB97,
      Spin;


Type
  TOM_DPMVTCAP = Class (TOM)
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnNewRecord  ; override ;
    procedure OnLoadRecord ; override ;
    procedure OnChangeField(F: TField);  override;
  private
    DPMSens     : String;
    EvovaleurNomin    : String;
  end ;


////////////// IMPLEMENTATION //////////////////
implementation

uses dpOutils,
     UDossierSelect,
     galOutil;

{ --------- Historique du capital ---------- }

procedure TOM_DPMVTCAP.OnArgument(stArgument: String);
var action : string;
begin
  Inherited ; // traite ACTION= sans l'enlever de stArgument
  action := ReadTokenSt (stArgument);
  DPMSens := ReadTokenSt (stArgument);
  if (DPMSens = '') then DPMSens := '+';
  EvovaleurNomin := ReadTokenSt (stArgument);
end;


procedure TOM_DPMVTCAP.OnLoadRecord;
var
  ValSens : THDBValCombobox ;
  ValEditNom  : THDBEdit;
  ValNominale : double;
begin
   ValNominale := valeur(EvovaleurNomin);
   ValEditNom := THDBEdit(GetControl('DPM_VALNOM'));
   if (DS.State in [dsInsert,dsEdit]) then
     begin
     ValEditNom.Text := EvovaleurNomin;
     SetField('DPM_VALNOM',ValNominale) ;
     SetControlEnabled('BFirst', False);
     SetControlEnabled('BPrev', False);
     SetControlEnabled('BNext', False);
     SetControlEnabled('BLast', False);
     end;
   ValSens := THDBValCombobox(GetControl('DPM_SENS'));
   if (ValSens<>Nil) AND (GetField('DPM_SENS')='') then
     ValSens.value := DPMSens;
end;


procedure TOM_DPMVTCAP.OnNewRecord;
begin
  SetField('DPM_GUIDPER', VH_Doss.GuidPer);

  SetField('DPM_NOORDRE',IncrementeSeqNoOrdre('DPM','DPMVTCAP',VH_Doss.GuidPer)) ;
end;


procedure TOM_DPMVTCAP.OnChangeField(F: TField);
var
  NbTitres : THDBSpinEdit;
  ValNominale : double;
  MontantTotal : double;
begin
  if (F.FieldName='DPM_NBTITRES') then
    begin
    NbTitres := THDBSpinEdit (GetControl('DPM_NBTITRES'));
    ValNominale := GetField ('DPM_VALNOM');
    if (ValNominale <> 0.00) then
      MontantTotal := NbTitres.Value * ValNominale
    else
      MontantTotal := 0;
    SetField('DPM_MONTANT',MontantTotal) ;
    end
end;



Initialization
  registerclasses([TOM_DPMVTCAP]) ;
end.
