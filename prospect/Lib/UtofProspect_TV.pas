unit UtofProspect_TV;

interface
uses  StdCtrls,Classes,forms,
      HEnt1,UTOF,UtilGC
      , UTobView, UtilSelection,UtilRT,ParamSoc
{$ifdef GIGI}
      ,EntRT
      ,UtofAfParamDp
      ,EntGC
{$ENDIF GIGI}
{$ifdef AFFAIRE}
      ,UtofAfTraducChampLibre
{$ENDIF}
{$IFDEF EAGLCLIENT}
     ,MaineAGL
{$ELSE}
     ,FE_Main
{$ENDIF}
      ;
Type
{$ifdef AFFAIRE}
                //mcd 24/11/2005 pour faire affectation depuis ressource si paramétré
     TOF_PROSPECTS_TV = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_PROSPECTS_TV = Class (TOF)
{$endif}
     private
         TobViewer1: TTobViewer;
         stProduitpgi : string;
         procedure TVOnDblClickCell(Sender: TObject ) ;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
     END ;

Procedure RTLanceFiche_PROSPECTS_TV (Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation

Procedure RTLanceFiche_PROSPECTS_TV(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_PROSPECTS_TV.OnArgument(Arguments : String ) ;
var F : TForm;
    Memo : TMemo;
{$ifdef GIGI}
    ListeChamp, ListeTable,Sql : string;
{$ENDIF GIGI}
begin
inherited ;
F := TForm (Ecran);
stProduitpgi := Arguments;
if stProduitpgi = '' then stProduitpgi := 'GRC';
if Arguments = 'GRF' then
  begin
  if GetParamSocSecur('SO_RTGESTINFOS003',False) = True then
      MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
  end
else MulCreerPagesCL(F,'NOMFIC=GCTIERS');
TobViewer1:=TTobViewer(getcontrol('TV'));
TobViewer1.OnDblClick:= TVOnDblClickCell ;
memo := TMemo(GetControl('FSQL'));
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
 if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('TSRESSOURCE',false)
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    Memo.lines[4]:='YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3,YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,';
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
 if (GetControl('TYTC_REPRESENTANT2') <> nil) then  SetControlVisible('TYTC_REPRESENTANT2',false);
 if (GetControl('YTC_REPRESENTANT3') <> nil) then  SetControlVisible('YTC_REPRESENTANT3',false);
 if (GetControl('TYTC_REPRESENTANT3') <> nil) then  SetControlVisible('TYTC_REPRESENTANT3',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
 SetControlProperty ('T_NATUREAUXI_', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI_', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI_', 'Plus', VH_GC.AfNatTiersGRCGI);
 Memo.lines[2]:='T_SECTEUR,T_SOCIETEGROUPE,T_PRESCRIPTEUR,T_JURIDIQUE,';
 Memo.lines[3]:='T_ORIGINETIERS,T_MOISCLOTURE,T_ENSEIGNE';  //MCD 04/07/2005
 If Vh_RT.TobChampsDpMul.detail.count <>0 then
   begin   //MCD 04/07/2005
   if (ecran).name='RTPROSPECTS_TV' then
     begin
     MulCreerPagesDP(F);
     ListeChamp :='';
     ListeTable :='';
     AfAjoutTableDp (ListeChamp,Listetable,'DP' ) ;
     Sql :='SELECT  T_NATUREAUXI, T_TIERS, T_LIBELLE, T_DATEPROCLI,'+
              ' T_CODEPOSTAL, T_VILLE, T_TELEPHONE, T_AUXILIAIRE, T_APE,'+
              ' T_SECTEUR,T_SOCIETEGROUPE,T_PRESCRIPTEUR,T_JURIDIQUE, '+
              ' T_ORIGINETIERS,T_MOISCLOTURE,T_ENSEIGNE,'+
              ' YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,'+
              ' YTC_TABLELIBRETIERS3,YTC_TABLELIBRETIERS4,YTC_TABLELIBRETIERS5,'+
              ' YTC_TABLELIBRETIERS6,YTC_TABLELIBRETIERS7,YTC_TABLELIBRETIERS8,'+
              ' YTC_TABLELIBRETIERS9,YTC_TABLELIBRETIERSA,'+
              ' PROSPECTS.*  ';
     Sql := SQL + listeChamp;
     Sql := Sql +' FROM TIERS LEFT OUTER JOIN PROSPECTS ON T_AUXILIAIRE=RPR_AUXILIAIRE '+
              ' LEFT OUTER JOIN TIERSCOMPL ON'+
              ' T_TIERS=YTC_TIERS ' + listeTable;
     Memo.text:=Sql;
     end
  end;
{$endif}
end;

procedure TOF_PROSPECTS_TV.OnLoad;
var Confid : string;
begin
inherited;
  if stProduitpgi = 'GRC' then Confid:='CON' else Confid:='CONF';
  SetControlText('XX_WHERE',RTXXWhereConfident(Confid,true)) ;
end;

procedure TOF_PROSPECTS_TV.TVOnDblClickCell(Sender: TObject );
begin
with TTobViewer(sender) do
  if stProduitpgi = 'GRF' then AGLLanceFiche('GC','GCFOURNISSEUR','',AsString[ColIndex('T_AUXILIAIRE'), CurrentRow],'MONOFICHE;T_NATUREAUXI='+AsString[ColIndex('T_NATUREAUXI'), CurrentRow])
  else AGLLanceFiche('GC','GCTIERS','',AsString[ColIndex('T_AUXILIAIRE'), CurrentRow],'MONOFICHE;T_NATUREAUXI='+AsString[ColIndex('T_NATUREAUXI'), CurrentRow])
end;

Initialization
registerclasses([TOF_PROSPECTS_TV]);

end.
