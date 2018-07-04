unit UTofCubeProspect;

interface
uses {mng Hcompte,}HEnt1,
     ParamSoc,   UtilGC,
{$IFDEF EAGLCLIENT}
     MainEAGL,
{$ELSE}
     Fe_Main,
{$ENDIF}
 {$ifdef GIGI}
      EntRt,EntGC,Cube,Hctrls,
 {$ENDIF GIGI}
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}
     Forms,UTOF, Classes
     ,UtilSelection,UtilRT;

const
     TableCombo : string = 'RTRPRLIBTABLE';
     TableComboMulti : string = 'RTRPRLIBTABMUL';
     ChampMul : string = 'RPR_RPRLIBMUL';
     PosMul : Integer = 14;
     HauteurPanel : integer = 30;
     LargeurPanel : integer = 120;
     LargeurLabel : integer = 163;
     NbPanelMax : integer = 15;
     ChampLeft : integer = 180;
     LabelLeft : integer = 15;
Type

{$ifdef AFFAIRE}
                //mcd 24/11/2005 pour faire affectation depuis ressource si paramétré
     TOF_CubeProspect = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_CubeProspect = Class (TOF)
{$endif}
     Private
        stProduitpgi : string;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
     private

    END ;

Procedure RTLanceFiche_CubeProspect (Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation

Procedure RTLanceFiche_CubeProspect(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_CubeProspect.OnArgument(Arguments : String ) ;
var F : TForm;
begin
inherited ;
  F := TForm (Ecran);
  if Arguments = 'GRF' then
    begin
    if GetParamSocSecur('SO_RTGESTINFOS003',False) = True then
        MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
    end
  else MulCreerPagesCL(F,'NOMFIC=GCTIERS');
  stProduitpgi := Arguments;
  if stProduitpgi = '' then stProduitpgi := 'GRC';
  if GetParamSocSecur('SO_RTGESTINFOS008',True) = True then
    MulCreerPagesCL(F,'NOMFIC=LIGNESPIECES');
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
 TFCUBE(Ecran).WhereSQL :=VH_GC.AfTypTiersGRCGI ;
 if Ecran.name ='RTTIERS_PIECE_CUB' then  Ecran.Caption := TraduireMemoire('Cube facturation (tiers/pièce)')
  else  Ecran.Caption := TraduireMemoire('Statistiques croisées sur tiers');
 UpdateCaption(Ecran);
 If Vh_RT.TobChampsDpMul.detail.count <>0 then
   begin //MCD 04/07/2005
   if (ecran).name='RTTIERS_CUBE' then
     begin
     MulCreerPagesDP(F);
     TFCUBE(Ecran).FromSQL :='TIERS left join PROSPECTS on T_AUXILIAIRE = RPR_AUXILIAIRE left join TIERSCOMPL ON T_TIERS=YTC_TIERS' +
                          ' left join annuaire on ann_tiers=t_tiers '+
                          ' left join annubis on ann_guidper=anb_guidper'+
                          ' left join dporga on ann_guidper=dor_guidper'+
                          ' left join dpsocial on ann_guidper=dso_guidper'+
                          ' left join dpfiscal on ann_guidper=dfi_guidper';
     end
   else if (ecran).name='RTTIERS_PIECE_CUB' then
       begin
       MulCreerPagesDP(F);
       TFCUBE(Ecran).FromSQL :=' TIERS LEFT JOIN PROSPECTS ON T_AUXILIAIRE = RPR_AUXILIAIRE'+
                          ' LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS '+
                          ' LEFT JOIN LIGNE ON GL_TIERS=T_TIERS left join RTINFOS008 on RD8_CLEDATA=RTRIM(GL_NATUREPIECEG)||";"||GL_SOUCHE||";"||CAST(GL_NUMERO AS VARCHAR(7))||";"||CAST(GL_INDICEG AS VARCHAR(7))||";"||CAST(GL_NUMLIGNE AS VARCHAR(7))'+
                          ' left join annuaire on ann_tiers=t_tiers '+
                          ' left join annubis on ann_guidper=anb_guidper'+
                          ' left join dporga on ann_guidper=dor_guidper'+
                          ' left join dpsocial on ann_guidper=dso_guidper'+
                          ' left join dpfiscal on ann_guidper=dfi_guidper';
       end;

   end;
{$endif}
end;

procedure TOF_CubeProspect.OnLoad;
var Confid : string;
begin
inherited;
  if stProduitpgi = 'GRC' then Confid:='CON' else Confid:='CONF';
  SetControlText('XX_WHERE',RTXXWhereConfident(Confid)) ;
end;

Initialization
registerclasses([TOF_CubeProspect]);
end.
