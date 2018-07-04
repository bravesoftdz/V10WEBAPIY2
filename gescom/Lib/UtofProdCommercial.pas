unit UtofProdCommercial;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,windows,vierge,Messages,
      HCtrls,HEnt1,HMsgBox,UTOF,Utob,M3FP,UtilSelection,entRT,comctrls,
{$IFDEF EAGLCLIENT}
     MaineAGL,eTablette,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,tablette,
{$ENDIF}
      UtilRT;

type
     TOF_RTcommercial_PRO = Class (TOF)
     private
        Deb_mois,Fin_Mois,Deb_Moins12,Fin_Moins12,Deb_Plus4,Fin_Plus4 :TDate;
        MontPrevuAntMoisP,MontCalAntMoisP: double;
        MontPrevuDuMoisP,MontCalDuMoisP : double;
        MontPrevuPostP,MontCalPostP : double;
        MontPrevuAntMoisC,MontCalAntMoisC: double;
        MontPrevuDuMoisC,MontCalDuMoisC : double;
        MontPrevuPostC,MontCalPostC : double;
        MTCdFor,MTCdPro :double;
        NbVisiteCli,NbVisitePro : integer;
        Q: TQuery;
        procedure ChargeInfosMois;
     public
        procedure OnLoad ; override ;
        procedure RTCalculRealise;
     END ;

implementation


procedure TOF_RTcommercial_PRO.OnLoad;
begin
Inherited;
if (VH_RT.RTCodeCommercial<>'') and (VH_RT.RTResponsable<>'') then
  RTCalculRealise;

 {PGIinfo('Aucun détail à afficher',Ecran.Caption); }
//PostMessage(TFVierge(Ecran).Handle,WM_CLOSE,0,0) ; exit; end;
End;

procedure TOF_RTcommercial_PRO.ChargeInfosMois;
begin
MontPrevuAntMoisP := 0.0;  MontCalAntMoisP := 0.0;
MontPrevuDuMoisP := 0.0; MontCalDuMoisP := 0.0;
MontPrevuPostP := 0.0;  MontCalPostP := 0.0;
MontPrevuAntMoisC := 0.0;  MontCalAntMoisC := 0.0;
MontPrevuDuMoisC := 0.0; MontCalDuMoisC := 0.0;
MontPrevuPostC := 0.0;  MontCalPostC := 0.0;
MTCdFor:= 0.0;  MTCdPro:= 0.0; NbVisiteCli:= 0; NbVisitePro:= 0;

Q:=OpenSQL(' SELECT sum(rpe_montantper) as MtPrevu,sum(rpe_montantper*rpe_pourcentage/100) as MtCalcule'+
           ' FROM perspectives '+
           ' where rpe_typeperspectiv="001" '+
           ' and rpe_daterealise >= "'+USDateTime(Deb_Moins12)+'" '+
           ' and rpe_daterealise <= "'+USDateTime(Fin_Moins12)+'" '+
           ' and (rpe_etatper="ENC" or rpe_etatper="DIF") '+
           ' and rpe_intervenant="'+VH_RT.RTResponsable+'"',true);

if not Q.EOF then
  begin
  MontPrevuAntMoisP := Q.FindField('MtPrevu').AsFloat;
  MontCalAntMoisP := Q.FindField('MtCalcule').AsFloat;
  end;
Ferme(Q);

Q:=OpenSQL(' SELECT sum(rpe_montantper) as MtPrevu,sum(rpe_montantper*rpe_pourcentage/100) as MtCalcule'+
           ' FROM perspectives '+
           ' where rpe_typeperspectiv="005" '+
           ' and rpe_daterealise >= "'+USDateTime(Deb_Moins12)+'" '+
           ' and rpe_daterealise <= "'+USDateTime(Fin_Moins12)+'" '+
           ' and (rpe_etatper="ENC" or rpe_etatper="DIF") '+
           ' and rpe_intervenant="'+VH_RT.RTResponsable+'"',true);

if not Q.EOF then
  begin
  MontPrevuAntMoisC := Q.FindField('MtPrevu').AsFloat;
  MontCalAntMoisC := Q.FindField('MtCalcule').AsFloat;
  end;
Ferme(Q);

Q:=OpenSQL(' SELECT sum(rpe_montantper) as MtPrevu,sum(rpe_montantper*rpe_pourcentage/100) as MtCalcule'+
           ' FROM perspectives '+
           ' where rpe_typeperspectiv="001" '+
           ' and rpe_daterealise >= "'+USDateTime(Deb_Mois)+'" '+
           ' and rpe_daterealise <= "'+USDateTime(Fin_Mois)+'" '+
           ' and (rpe_etatper="ENC" or rpe_etatper="DIF") '+
           ' and rpe_intervenant="'+VH_RT.RTResponsable+'"',true);

if not Q.EOF then
  begin
  MontPrevuDuMoisP := Q.FindField('MtPrevu').AsFloat;
  MontCalDuMoisP := Q.FindField('MtCalcule').AsFloat;
  end;
Ferme(Q);

Q:=OpenSQL(' SELECT sum(rpe_montantper) as MtPrevu,sum(rpe_montantper*rpe_pourcentage/100) as MtCalcule'+
           ' FROM perspectives '+
           ' where rpe_typeperspectiv="005" '+
           ' and rpe_daterealise >= "'+USDateTime(Deb_Mois)+'" '+
           ' and rpe_daterealise <= "'+USDateTime(Fin_Mois)+'" '+
           ' and (rpe_etatper="ENC" or rpe_etatper="DIF") '+
           ' and rpe_intervenant="'+VH_RT.RTResponsable+'"',true);

if not Q.EOF then
  begin
  MontPrevuDuMoisC := Q.FindField('MtPrevu').AsFloat;
  MontCalDuMoisC := Q.FindField('MtCalcule').AsFloat;
  end;
Ferme(Q);

Q:=OpenSQL(' SELECT sum(rpe_montantper) as MtPrevu,sum(rpe_montantper*rpe_pourcentage/100) as MtCalcule'+
           ' FROM perspectives '+
           ' where rpe_typeperspectiv="001" '+
           ' and rpe_daterealise >= "'+USDateTime(Deb_Plus4)+'" '+
           ' and rpe_daterealise <= "'+USDateTime(Fin_Plus4)+'" '+
           ' and (rpe_etatper="ENC" or rpe_etatper="DIF") '+
           ' and rpe_intervenant="'+VH_RT.RTResponsable+'"',true);

if not Q.EOF then
  begin
  MontPrevuPostP := Q.FindField('MtPrevu').AsFloat;
  MontCalPostP := Q.FindField('MtCalcule').AsFloat;
  end;
Ferme(Q);

Q:=OpenSQL(' SELECT sum(rpe_montantper) as MtPrevu,sum(rpe_montantper*rpe_pourcentage/100) as MtCalcule'+
           ' FROM perspectives '+
           ' where rpe_typeperspectiv="005" '+
           ' and rpe_daterealise >= "'+USDateTime(Deb_Plus4)+'" '+
           ' and rpe_daterealise <= "'+USDateTime(Fin_Plus4)+'" '+
           ' and (rpe_etatper="ENC" or rpe_etatper="DIF") '+
           ' and rpe_intervenant="'+VH_RT.RTResponsable+'"',true);

if not Q.EOF then
  begin
  MontPrevuPostC := Q.FindField('MtPrevu').AsFloat;
  MontCalPostC := Q.FindField('MtCalcule').AsFloat;
  end;
Ferme(Q);

Q:=OpenSQL(' SELECT count(rac_numaction) as nb_visite_c'+
           ' FROM actions '+
           ' left join tiers on rac_tiers=t_tiers '+
           ' WHERE t_natureauxi="CLI" '+
           ' and rac_typeaction="001" '+
           ' and rac_etataction="REA" '+
           ' and rac_dateaction >= "'+USDateTime(Deb_Mois)+'" '+
           ' and rac_dateaction <= "'+USDateTime(Fin_Mois)+'" '+
           ' and rac_intervenant="'+VH_RT.RTResponsable+'"',true);

if not Q.EOF then NbVisiteCli := Q.FindField('nb_visite_c').AsInteger;
Ferme(Q);

Q:=OpenSQL(' SELECT count(rac_numaction) as nb_visite_p'+
           ' FROM actions '+
           ' left join tiers on rac_tiers=t_tiers '+
           ' WHERE t_natureauxi="PRO" '+
           ' and rac_typeaction="001" '+
           ' and rac_etataction="REA" '+
           ' and rac_dateaction >= "'+USDateTime(Deb_Mois)+'" '+
           ' and rac_dateaction <= "'+USDateTime(Fin_Mois)+'" '+
           ' and rac_intervenant="'+VH_RT.RTResponsable+'"',true);

if not Q.EOF then NbVisitePro := Q.FindField('nb_visite_p').AsInteger;
Ferme(Q);

Q:=OpenSQL(' SELECT sum(gp_totalht) as mnt_cde_f'+
           ' FROM piece'+
           ' where gp_naturepieceg="CC" '+
           ' and gp_datepiece >= "'+USDateTime(Deb_Mois)+'" '+
           ' and gp_datepiece <= "'+USDateTime(Fin_Mois)+'" '+
           ' and gp_representant="'+VH_RT.RTCodeCommercial+'"'+
           ' and gp_domaine = "003"',true);

if not Q.EOF then MTCdFor := Q.FindField('mnt_cde_f').AsFloat;
Ferme(Q);

Q:=OpenSQL(' SELECT sum(gp_totalht) as mnt_cde_p'+
           ' FROM piece'+
           ' where gp_naturepieceg="CC"'+
           ' and gp_datepiece >= "'+USDateTime(Deb_Mois)+'" '+
           ' and gp_datepiece <= "'+USDateTime(Fin_Mois)+'" '+
           ' and gp_representant="'+VH_RT.RTCodeCommercial+'"'+
           ' and gp_domaine = "004"',true);

if not Q.EOF then MTCdPro := Q.FindField('mnt_cde_p').AsFloat;
Ferme(Q);
end;

procedure TOF_RTcommercial_PRO.RTCalculRealise;
var    TabSheet: TTabSheet;
begin
SetControlText('TGCL_LIBELLE', VH_RT.RTNomCommercial);

Deb_Mois := DebutDeMois (V_PGI.DateEntree);
Fin_Mois := FinDeMois (V_PGI.DateEntree);
Fin_Moins12 := FinDeMois  (PlusMois (Fin_Mois, -1));
Deb_Moins12 := DebutDeMois(PlusMois (Fin_Moins12,-11));
Deb_Plus4 := Deb_Mois;
Fin_Plus4 := FinDeMois(PlusMois (Deb_Plus4,3));

if (Transactions (ChargeInfosMois,1)=oeOk) then
begin
  TabSheet := TTabSheet(Ecran.FindComponent('PMOISCOURS'));
  TPageControl(getcontrol('PCONTROL')).activePage := TabSheet;
  TabSheet.caption := FormatDateTime('mmmm', Deb_Mois );
  SetControlText('DATEPROPDEB', DateToStr (Deb_Mois) );
  SetControlText('DATEPROPFIN', DateToStr (Fin_Mois) );

  THNumEdit(GetControl('NBVISITEPRO')).Masks.PositiveMask :='#0';
  THNumEdit(GetControl('NBVISITECLI')).Masks.PositiveMask :='#0';
  THNumEdit(GetControl('NBVISITETOT')).Masks.PositiveMask :='#0';

  THNumEdit (Getcontrol('MTPREANTP')).Value := MontPrevuAntMoisP;
  THNumEdit (Getcontrol('MTCALANTP')).Value := MontCalAntMoisP;
  THNumEdit (Getcontrol('MTPREDUMOISP')).Value := MontPrevuDuMoisP;
  THNumEdit (Getcontrol('MTCALDUMOISP')).Value := MontCalDuMoisP;
  THNumEdit (Getcontrol('MTPREPOSTP')).Value := MontPrevuPostP;
  THNumEdit (Getcontrol('MTCALPOSTP')).Value := MontCalPostP;

  THNumEdit (Getcontrol('MTPREANTC')).Value := MontPrevuAntMoisC;
  THNumEdit (Getcontrol('MTCALANTC')).Value := MontCalAntMoisC;
  THNumEdit (Getcontrol('MTPREDUMOISC')).Value := MontPrevuDuMoisC;
  THNumEdit (Getcontrol('MTCALDUMOISC')).Value := MontCalDuMoisC;
  THNumEdit (Getcontrol('MTPREPOSTC')).Value := MontPrevuPostC;
  THNumEdit (Getcontrol('MTCALPOSTC')).Value := MontCalPostC;

  THNumEdit (Getcontrol('MTCDPRO')).Value := MTCdPro;
  THNumEdit (Getcontrol('MTCDFOR')).Value := MTCdFor;
  THNumEdit (Getcontrol('MTCDTOT')).Value := MTCdPro+MTCdFor;

  THNumEdit (Getcontrol('NBVISITEPRO')).Value := NbVisitePro;
  THNumEdit (Getcontrol('NBVISITECLI')).Value := NbVisiteCli;
  THNumEdit (Getcontrol('NBVISITETOT')).Value := NbVisitePro+NbVisiteCli;
end;

Deb_Mois := DebutDeMois ((PlusMois (V_PGI.DateEntree, -1)));
Fin_Mois := FinDeMois   (Deb_Mois);
Fin_Moins12 := FinDeMois  (PlusMois (Fin_Mois, -1));
Deb_Moins12 := DebutDeMois(PlusMois (Fin_Moins12,-11));
Deb_Plus4 := Deb_Mois;
Fin_Plus4 := FinDeMois(PlusMois (Deb_Plus4,3));

if (Transactions (ChargeInfosMois,1)=oeOk) then
begin
//  TPageControl(getcontrol('PCONTROL')).activePage := TTabSheet(Ecran.FindComponent('PMOISPREC'));
  TabSheet := TTabSheet(Ecran.FindComponent('PMOISPREC'));
  TPageControl(getcontrol('PCONTROL')).activePage := TabSheet;
  TabSheet.caption := FormatDateTime('mmmm', Deb_Mois );

  SetControlText('DATEPROPDEB1', DateToStr (Deb_Mois) );
  SetControlText('DATEPROPFIN1', DateToStr (Fin_Mois) );

  THNumEdit(GetControl('NBVISITEPRO1')).Masks.PositiveMask :='#0';
  THNumEdit(GetControl('NBVISITECLI1')).Masks.PositiveMask :='#0';
  THNumEdit(GetControl('NBVISITETOT1')).Masks.PositiveMask :='#0';

  THNumEdit (Getcontrol('MTPREANTP1')).Value := MontPrevuAntMoisP;
  THNumEdit (Getcontrol('MTCALANTP1')).Value := MontCalAntMoisP;
  THNumEdit (Getcontrol('MTPREDUMOISP1')).Value := MontPrevuDuMoisP;
  THNumEdit (Getcontrol('MTCALDUMOISP1')).Value := MontCalDuMoisP;
  THNumEdit (Getcontrol('MTPREPOSTP1')).Value := MontPrevuPostP;
  THNumEdit (Getcontrol('MTCALPOSTP1')).Value := MontCalPostP;

  THNumEdit (Getcontrol('MTPREANTC1')).Value := MontPrevuAntMoisC;
  THNumEdit (Getcontrol('MTCALANTC1')).Value := MontCalAntMoisC;
  THNumEdit (Getcontrol('MTPREDUMOISC1')).Value := MontPrevuDuMoisC;
  THNumEdit (Getcontrol('MTCALDUMOISC1')).Value := MontCalDuMoisC;
  THNumEdit (Getcontrol('MTPREPOSTC1')).Value := MontPrevuPostC;
  THNumEdit (Getcontrol('MTCALPOSTC1')).Value := MontCalPostC;

  THNumEdit (Getcontrol('MTCDPRO1')).Value := MTCdPro;
  THNumEdit (Getcontrol('MTCDFOR1')).Value := MTCdFor;
  THNumEdit (Getcontrol('MTCDTOT1')).Value := MTCdPro+MTCdFor;

  THNumEdit (Getcontrol('NBVISITEPRO1')).Value := NbVisitePro;
  THNumEdit (Getcontrol('NBVISITECLI1')).Value := NbVisiteCli;
  THNumEdit (Getcontrol('NBVISITETOT1')).Value := NbVisitePro+NbVisiteCli;
end;

TPageControl(getcontrol('PCONTROL')).activePage := TTabSheet(Ecran.FindComponent('PMOISCOURS'));

end;

Initialization
registerclasses([TOF_RTcommercial_PRO]);

end.
