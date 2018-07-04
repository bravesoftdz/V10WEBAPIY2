unit UtofActions_TV;

interface
uses  Classes,forms,sysutils,
      HEnt1,UTOF, UTobView,UtilGc ,StdCtrls
{$IFDEF EAGLCLIENT}
      ,MainEAGL
{$ELSE}
      ,Fe_Main
{$ENDIF}
{$ifdef AFFAIRE}
      ,UtofAfTraducChampLibre
{$ENDIF}
      ,ParamSoc,UtilSelection,UtilRT,utofAfBaseCodeAffaire,hctrls,EntGC,UtilPGI,HMsgBox;

Function RTLanceFiche_Actions_TV(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
{$ifdef AFFAIRE}
                //mcd 11/05/2006 12940  pour faire affectation depuis ressource si paramétré
     TOF_Actions_TV = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
    TOF_Actions_TV = Class (TOF_AFBASECODEAFFAIRE)
{$endif}
{$IFDEF AFFAIRE}
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override;
        procedure bSelectAff1Click(Sender: TObject);     override ;
{$ENDIF AFFAIRE}
     private
         TobViewer1: TTobViewer;
         stProduitpgi : string;
         procedure ActTVOnDblClickCell(Sender: TObject ) ;
         function EstMultiSocEnreg(Sender: TObject) : boolean;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
     END ;



implementation

uses entRT;
Function RTLanceFiche_Actions_TV(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;


procedure TOF_Actions_TV.OnArgument(Arguments : String ) ;
var F : TForm;
    Memo : TMemo;
begin
	fMulDeTraitement := true;
inherited ;
  if Arguments <> 'GRF' then
    begin
    F := TForm (Ecran);
    MulCreerPagesCL(F,'NOMFIC=GCTIERS');

    if GetParamSocSecur('SO_RTGESTINFOS001',False) = True then
        MulCreerPagesCL(F,'NOMFIC=RTACTIONS');
    stProduitpgi := 'GRC';
    end
  else stProduitpgi := Arguments;
  TobViewer1:=TTobViewer(getcontrol('TV'));
  {if (GetControl('MULTIDOSSIER') = nil) then }TobViewer1.OnDblClick:= ActTVOnDblClickCell ;
  if (Ecran.name = 'RTACTIONS_TVMTIE') then
  begin
  setcontroltext   ('MULTIDOSSIER',MS_CODEREGROUPEMENT);
  end;
{$IFDEF AFFAIRE}
     if ( not (ctxAffaire in V_PGI.PGIContexte) ) and ( not ( ctxGCAFF in V_PGI.PGIContexte) ) then
{$ENDIF}     
      begin
      SetControlVisible ('BEFFACEAFF1',false); SetControlVisible ('BSELECTAFF1',false);
      SetControlVisible ('TRAC_AFFAIRE',false); SetControlVisible ('RAC_AFFAIRE1',false);
      SetControlVisible ('RAC_AFFAIRE2',false); SetControlVisible ('RAC_AFFAIRE3',false);
      SetControlVisible ('RAC_AVENANT',false);
      end;
memo := TMemo(GetControl('FSQL'));
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    Memo.lines[12]:=',YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3,YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,YTC_TABLELIBRETIERS3';
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
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
 Memo.lines[11]:=',T_PRESCRIPTEUR,T_SOCIETEGROUPE';
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

procedure TOF_Actions_TV.OnLoad;
var Confid : string;
begin
inherited;
  if stProduitpgi = 'GRC' then Confid:='CON' else Confid:='CONF';
  SetControlText('XX_WHERE',RTXXWhereConfident(Confid)) ;
end;

procedure TOF_Actions_TV.ActTVOnDblClickCell(Sender: TObject );
var Stchaine : string;
    stAction   : TActionFiche ;
begin
with TTobViewer(sender) do
    begin
    if  (EstMultiSocEnreg (sender) = False) then
        begin
        stAction:=taConsult;
    if (ColName[CurrentCol] = 'RAC_AUXILIAIRE') or (ColName[CurrentCol] = 'T_LIBELLE') or (ColName[CurrentCol] = 'RAC_TIERS')then
      begin
      if stProduitPGI = 'GRC' then
        V_PGI.DispatchTT (28,taConsult ,AsString[ColIndex('RAC_AUXILIAIRE'), CurrentRow], '','')
      else
        V_PGI.DispatchTT (29,taConsult ,AsString[ColIndex('RAC_AUXILIAIRE'), CurrentRow], '','')
      end
    else if (ColName[CurrentCol] = 'RAC_INTERVENANT') or (ColName[CurrentCol] = 'ARS_LIBELLE') then
//      V_PGI.DispatchTT (9,taConsult ,AsString[ColIndex('RAC_INTERVENANT'), CurrentRow], '','')
      V_PGI.DispatchTT (6,taConsult ,AsString[ColIndex('RAC_INTERVENANT'), CurrentRow], '','')
    else if (ColName[CurrentCol] = 'RAC_PROJET') then
        begin
        StChaine := AsString[ColIndex('RAC_PROJET'), CurrentRow];
        if Stchaine <> '' then
           V_PGI.DispatchTT (30,taConsult ,AsString[ColIndex('RAC_PROJET'), CurrentRow], '','')
        end
    else if (ColName[CurrentCol] = 'RAC_OPERATION') then
        begin
        StChaine := AsString[ColIndex('RAC_OPERATION'), CurrentRow];
        if Stchaine <> '' then
           V_PGI.DispatchTT (23,taConsult ,AsString[ColIndex('RAC_OPERATION'), CurrentRow], '','')
        end
    else
        begin
        if copy(ColName[CurrentCol],1,3) = 'RCH' then
           V_PGI.DispatchTT (24,taConsult ,IntToStr(AsInteger[ColIndex('RCH_NUMERO'), CurrentRow]), '','PRODUITPGI='+stProduitPGI)
        else
           begin
           if ( (stProduitPGI = 'GRC') and
              ( ( (VH_RT.RTCreatActions) or (RTDroitModifTiers(AsString[ColIndex('RAC_TIERS'), CurrentRow]) ) ) and
                (RTDroitModifActions('',AsString[ColIndex('RAC_TYPEACTION'), CurrentRow],AsString[ColIndex('RAC_INTERVENANT'), CurrentRow]))) ) or
              ( (stProduitPGI = 'GRF') and
              ( ( (VH_RT.RFCreatActions) or (RTDroitModifFou(AsString[ColIndex('RAC_TIERS'), CurrentRow]) ) ) and
              (RTDroitModifActionsF('',AsString[ColIndex('RAC_TYPEACTION'), CurrentRow],AsString[ColIndex('RAC_INTERVENANT'), CurrentRow]))) )
              then stAction:=taModif;
           V_PGI.DispatchTT (22,stAction ,AsString[ColIndex('RAC_AUXILIAIRE'), CurrentRow]+';'+IntToStr(AsInteger[ColIndex('RAC_NUMACTION'), CurrentRow]), '',';NOCREAT;PRODUITPGI='+stProduitPGI);
           end;
            end;
        end;
    end;
end;

{$IFDEF AFFAIRE}
procedure TOF_Actions_TV.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('RAC_AFFAIRE'));
Aff1:=THEdit(GetControl('RAC_AFFAIRE1'));
Aff2:=THEdit(GetControl('RAC_AFFAIRE2'));
Aff3:=THEdit(GetControl('RAC_AFFAIRE3'));
Aff4:=THEdit(GetControl('RAC_AVENANT'));
Tiers:=THEdit(GetControl('RAC_TIERS'));
end;

procedure TOF_Actions_TV.bSelectAff1Click(Sender: TObject);
begin
    SelectionAffaire (EditTiers, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4, VH_GC.GASeria , false, '', false, true, true)
end;
{$ENDIF AFFAIRE}

function TOF_Actions_TV.EstMultiSocEnreg(Sender: TObject) : boolean;
begin
  with TTobViewer(sender) do
  begin
    Result := false;
    if Assigned (GetControl ('MULTIDOSSIER')) then
    begin
      if THValComboBox(GetControl('MULTIDOSSIER')).Value <> '' then
      begin
        if AsString[ColIndex('SYSDOSSIER'), CurrentRow] <> V_PGI.SchemaName then
        begin
          PgiBox ('Consultation impossible : les données ne se trouvent pas dans la base en cours');
          Result := True;
        end;
      end;
    end;
  end;
end;

Initialization
registerclasses([TOF_Actions_TV]);

end.
