unit UtofCubeActions;

interface

uses  Classes,forms,
      HCtrls,HEnt1,UTOF,UtilGc,
{$IFDEF EAGLCLIENT}
     MainEAGL,
{$ELSE}
     Fe_Main,
{$ENDIF}
      UtilSelection,UtilRT,ParamSoc,utofAfBaseCodeAffaire,EntGC;

Type
    TOF_CubeActions = Class (TOF_AFBASECODEAFFAIRE)
{$IFDEF AFFAIRE}
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override;
        procedure bSelectAff1Click(Sender: TObject);     override ;
{$ENDIF AFFAIRE}
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnLoad ; override ;
     private
         stProduitpgi : string;
     END;

Function RTLanceFiche_CubeActions(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation

Function RTLanceFiche_CubeActions(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_CubeActions.OnArgument(stArgument : String );
var F : TForm;
begin
inherited ;
  if stArgument <> 'GRF' then
    begin
    F := TForm (Ecran);
    MulCreerPagesCL(F,'NOMFIC=GCTIERS');
    if GetParamSocSecur('SO_RTGESTINFOS001',False) = True then
        MulCreerPagesCL(F,'NOMFIC=RTACTIONS');
    stProduitpgi := 'GRC';
    end
    else stProduitpgi := stArgument;
{$IFDEF AFFAIRE}
if ( not (ctxAffaire in V_PGI.PGIContexte) ) and ( not ( ctxGCAFF in V_PGI.PGIContexte) ) then
{$ENDIF}
      begin
      SetControlVisible ('BEFFACEAFF1',false); SetControlVisible ('BSELECTAFF1',false);
      SetControlVisible ('TRAC_AFFAIRE',false); SetControlVisible ('RAC_AFFAIRE1',false);
      SetControlVisible ('RAC_AFFAIRE2',false); SetControlVisible ('RAC_AFFAIRE3',false);
      SetControlVisible ('RAC_AVENANT',false);
      end;
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then
    begin
    SetControlVisible ('YTC_RESSOURCE1',false);
    SetControlVisible ('YTC_RESSOURCE2',false);
    SetControlVisible ('YTC_RESSOURCE3',false);
    SetControlVisible ('TYTC_RESSOURCE1',false);
    SetControlVisible ('TYTC_RESSOURCE2',false);
    SetControlVisible ('TYTC_RESSOURCE3',false);
    SetControlVisible ('T_MOISCLOTURE',false);
    SetControlVisible ('T_MOISCLOTURE_',false);
    SetControlVisible ('TT_MOISCLOTURE',false);
    SetControlVisible ('TT_MOISCLOTURE_',false);
    end
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    if not (ctxscot in V_PGI.PGICOntexte) then
       begin
       SetControlVisible ('T_MOISCLOTURE',false);
       SetControlVisible ('T_MOISCLOTURE_',false);
       SetControlVisible ('TT_MOISCLOTURE',false);
       SetControlVisible ('TT_MOISCLOTURE_',false);
       end;
    end;
  end;
{$Ifdef GIGI}
 if (GetControl('RAC_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    end;
 if (GetControl('RAC_PROJET') <> nil) and Not GetParamsocSecur('SO_RTPROJGESTION',False)
    then  begin
    SetControlVisible('RAC_PROJET',false);
    SetControlVisible('TRAC_PROJET',false);
    end;
 If (Not VH_GC.GaSeria) or not (GetParamSocSecur ('SO_AFRTPROPOS',False)) then
   begin
   if (GetControl('TRAC_AFFAIRE') <> nil) then  SetControlVisible('TRAC_AFFAIRE',false);
   if (GetControl('RAC_AFFAIRE1') <> nil) then  SetControlVisible('RAC_AFFAIRE1',false);
   if (GetControl('RAC_AFFAIRE2') <> nil) then  SetControlVisible('RAC_AFFAIRE2',false);
   if (GetControl('RAC_AFFAIRE3') <> nil) then  SetControlVisible('RAC_AFFAIRE3',false);
   if (GetControl('RAC_AVENANT') <> nil) then  SetControlVisible('RAC_AVENANT',false);
   if (GetControl('BEFFACEAFF1') <> nil) then  SetControlVisible('BEFFACEAFF1',false);
   if (GetControl('BSELECTAFF1') <> nil) then  SetControlVisible('BSELECTAFF1',false);
   end;
 if (GetControl('T_REPRESENTANT') <> nil) then  SetControlVisible('T_REPRESENTANT',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then  SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('YTC_REPRESENTANT2') <> nil) then  SetControlVisible('YTC_REPRESENTANT2',false);
 if (GetControl('YTC_REPRESENTANT2_') <> nil) then  SetControlVisible('YTC_REPRESENTANT2_',false);
 if (GetControl('YTC_REPRESENTANT3') <> nil) then  SetControlVisible('YTC_REPRESENTANT3',false);
 if (GetControl('YTC_REPRESENTANT3_') <> nil) then  SetControlVisible('YTC_REPRESENTANT3_',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
{$endif}
{$IFDEF GRCLIGHT}
  if (stProduitpgi = 'GRC') and (not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) ) then
    begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    SetControlVisible('RAC_PROJET',false);
    SetControlVisible('TRAC_PROJET',false);
    end;
{$ENDIF GRCLIGHT}

end;

procedure TOF_CubeActions.OnLoad;
var Confid : string;
begin
  inherited;
  if stProduitpgi = 'GRC' then Confid:='CON' else Confid:='CONF';
  SetControlText('XX_WHERE',RTXXWhereConfident(Confid)) ;
end;

{$IFDEF AFFAIRE}
procedure TOF_CubeActions.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('RAC_AFFAIRE'));
Aff1:=THEdit(GetControl('RAC_AFFAIRE1'));
Aff2:=THEdit(GetControl('RAC_AFFAIRE2'));
Aff3:=THEdit(GetControl('RAC_AFFAIRE3'));
Aff4:=THEdit(GetControl('RAC_AVENANT'));
Tiers:=THEdit(GetControl('RAC_TIERS'));
end;

procedure TOF_CubeActions.bSelectAff1Click(Sender: TObject);
begin
{$IFNDEF BTP}
    SelectionAffaire (EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4, VH_GC.GASeria , false, '', false, true, true)
{$ELSE}
    SelectionAffaire (EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4)
{$ENDIF}
end;
{$ENDIF AFFAIRE}

Initialization
registerclasses([TOF_CubeActions]);

end.
