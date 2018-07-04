unit UtofPros_ssactions;

interface
uses
  StdCtrls, Classes, Forms, SysUtils, Windows,
  {$IFDEF EAGLCLIENT}
    MaineAGL,
  {$ELSE}
    Fe_Main,
  {$ENDIF}
  HCtrls, HEnt1, UTof,
  {$IFDEF AFFAIRE}
    uTofAfTraducChampLibre,
  {$ENDIF}
  UTobView
  ;

Type
  {$IFDEF AFFAIRE}
                //mcd 24/11/2005 pour faire affectation depuis ressource si paramétré
     TOF_PROS_SSACTIONS = Class (TOF_AFTRADUCCHAMPLIBRE)
  {$ELSE}
     TOF_PROS_SSACTIONS = Class (TOF)
  {$ENDIF}
  private
   TobViewer1: TTobViewer;
   stProduitpgi : string;
   procedure TVOnDblClickCell(Sender: TObject ) ;
  public
    procedure OnArgument(Arguments : String ) ; override ;
    procedure OnLoad ; override ;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  END;

Procedure RTLanceFiche_PROS_SSACTIONS (Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation

uses
  UtilGC, UtilSelection,
  {$IFDEF GIGI}
    EntRT,
    UtofAfParamDp,
    EntGC,
  {$ENDIF GIGI}
  HeureUtil,
  UtilRT, Stat,
  ParamSoc
  ;

Procedure RTLanceFiche_PROS_SSACTIONS(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_PROS_SSACTIONS.OnArgument(Arguments : String ) ;
var F : TForm;
    DateJour : TdateTime;
    Memo : TMemo;
{$ifdef GIGI}
    ListeChamp, ListeTable,Sql : string;
{$ENDIF GIGI}
begin
inherited ;
stProduitpgi := Arguments;
if stProduitpgi = '' then stProduitpgi := 'GRC';
F := TForm (Ecran);
if stProduitpgi = 'GRF' then
  begin
  if GetParamSocSecur('SO_RTGESTINFOS003',False) = True then
      MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
  end
else if (TForm(Ecran).name <> 'RTPROSSACT_ZOOM') then MulCreerPagesCL(F,'NOMFIC=GCTIERS');
TobViewer1:=TTobViewer(getcontrol('TV'));
TobViewer1.OnDblClick:= TVOnDblClickCell ;
if stProduitpgi <> 'GRF' then THValComboBox (Ecran.FindComponent('T_NATUREAUXI_')).ItemIndex := 1;
DateJour := CurrentDate;
SetControlText ('DATEACTION_',DateToStr(DateJour));
DateJour := PlusDate(DateJour,-1,'A');
SetControlText ('DATEACTION',DateToStr(DateJour));
if (ecran <> Nil)then TFStat(ecran).OnKeyDown:=FormKeyDown ;
memo := TMemo(GetControl('FSQL'));
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
 if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
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
 if (GetControl('OPERATIONACT') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('OPERATIONACT',false);
    SetControlVisible('TOPERATIONACT',false);
    end;
 if (GetControl('T_REPRESENTANT') <> nil) then  SetControlVisible('T_REPRESENTANT',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then  SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('YTC_REPRESENTANT2') <> nil) then  SetControlVisible('YTC_REPRESENTANT2',false);
 if (GetControl('TYTC_REPRESENTANT2') <> nil) then  SetControlVisible('TYTC_REPRESENTANT2',false);
 if (GetControl('YTC_REPRESENTANT3') <> nil) then  SetControlVisible('YTC_REPRESENTANT3',false);
 if (GetControl('TYTC_REPRESENTANT3') <> nil) then  SetControlVisible('TYTC_REPRESENTANT3',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
 SetControlProperty ('T_NATUREAUXI_', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI_', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI_', 'Plus', VH_GC.AfNatTiersGRCGI);
 Ecran.Caption := TraduireMemoire('Suivi tiers/actions');
 UpdateCaption(Ecran);
 Memo.lines[2]:='T_SECTEUR,T_SOCIETEGROUPE,T_PRESCRIPTEUR,T_JURIDIQUE,';
 Memo.lines[3]:='T_PAYS,T_ORIGINETIERS,T_MOISCLOTURE,T_ENSEIGNE,';  //MCD 04/07/2005
 If Vh_RT.TobChampsDpMul.detail.count <>0 then
   begin   //MCD 04/07/2005
   if (ecran).name='RTPROSSACTION_TV' then
     begin
     MulCreerPagesDP(F);
     ListeChamp :='';
     ListeTable :='';
     AfAjoutTableDp (ListeChamp,Listetable,'DP' ) ;
     Sql := 'SELECT  T_NATUREAUXI, T_TIERS, T_LIBELLE,'+ 
            'T_CODEPOSTAL, T_VILLE,  T_AUXILIAIRE, T_APE, '+
            'T_SECTEUR,T_SOCIETEGROUPE,T_PRESCRIPTEUR,T_JURIDIQUE,'+
            'T_PAYS,T_ORIGINETIERS,T_MOISCLOTURE,T_ENSEIGNE, '+
            'YTC_TABLELIBRETIERS1, YTC_TABLELIBRETIERS2,YTC_TABLELIBRETIERS3,'+
            'YTC_TABLELIBRETIERS4, YTC_TABLELIBRETIERS5,YTC_TABLELIBRETIERS6,'+
            'YTC_TABLELIBRETIERS7, YTC_TABLELIBRETIERS8,YTC_TABLELIBRETIERS9,'+
            'YTC_TABLELIBRETIERSA, '+
            'YTC_VALLIBRE1,YTC_VALLIBRE2,YTC_VALLIBRE3,'+
            'YTC_DATELIBRE1,YTC_DATELIBRE2,YTC_DATELIBRE3, '+
            'YTC_BOOLLIBRE1,YTC_BOOLLIBRE2,YTC_BOOLLIBRE3,'+
            'PROSPECTS.*';
     Sql := SQL + listeChamp;
     Sql := Sql +' FROM TIERS LEFT OUTER JOIN PROSPECTS ON T_AUXILIAIRE=RPR_AUXILIAIRE '+
               ' LEFT OUTER JOIN TIERSCOMPL ON T_AUXILIAIRE=YTC_AUXILIAIRE'
               + listeTable;
     Memo.text:=Sql;
     end
  end;
{$endif}
{$IFDEF GRCLIGHT}
  if ( stProduitpgi = 'GRC' ) and (not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False)) then
    begin
    SetControlVisible('OPERATIONACT',false);
    SetControlVisible('TOPERATIONACT',false);
    end;
{$ENDIF GRCLIGHT}

end;

procedure TOF_PROS_SSACTIONS.OnLoad;
var StWhere,StListeActions,StAction,StOr,ListeCombos,Confid : string;
    DateDeb,DateFin : TDateTime;
    i: integer;
begin
Inherited;
StListeActions := GetControlText ('TYPEACTION');
DateDeb := StrToDate(GetControlText('DATEACTION'));
DateFin := StrToDate(GetControlText('DATEACTION_'));
StWhere := '';
if GetControlText('SANS') = 'X' then StWhere := 'NOT ';
StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = TIERS.T_AUXILIAIRE';
StOr:='';
if StListeActions <> TraduireMemoire('<<Tous>>') then
   begin
   StWhere := StWhere + ' AND (';
   While StListeActions <> '' do
      begin
      StAction :=ReadTokenSt(StListeActions);
      StWhere := StWhere + StOr + 'RAC_TYPEACTION = "'+ StAction + '"';
      StOr := ' OR ';
      end;
   StWhere := StWhere + ')';
   end;
StWhere := StWhere + ' AND RAC_DATEACTION >= "'+UsDateTime(DateDeb) +'" AND RAC_DATEACTION <= "'+UsDateTime(DateFin)+'"';
if (TForm(Ecran).name <> 'RTPROSSACT_ZOOM') then
   begin
    for i:=1 to 3 do
        begin
        if (GetControlText('TABLELIBRE'+intToStr(i)) <> '') and (GetControlText('TABLELIBRE'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
            begin
            ListeCombos:=FindEtReplace(GetControlText('TABLELIBRE'+intToStr(i)),';','","',True);
            ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
            if stProduitpgi = 'GRF' then StWhere := StWhere + ' AND RAC_TABLELIBREF' +intToStr(i) +' in ' + ListeCombos
            else StWhere := StWhere + ' AND RAC_TABLELIBRE' +intToStr(i) +' in ' + ListeCombos;
            end;
        end;
    if GetControlText('OPERATIONACT') <> '' then
      begin
      ListeCombos := GetControlText('OPERATIONACT');
      if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
      ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
      ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
//      StWhere:=StWhere + ' AND RAC_OPERATION = "'+GetControlText('OPERATIONACT')+'"';
      StWhere:=StWhere + ' AND RAC_OPERATION in '+ListeCombos;
      end;
    if (GetControlText('ETATACTION') <> '') and (GetControlText('ETATACTION') <> TraduireMemoire('<<Tous>>')) then
        begin
        ListeCombos:=FindEtReplace(GetControlText('ETATACTION'),';','","',True);
        ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
        StWhere := StWhere + ' AND RAC_ETATACTION in ' + ListeCombos;
        end;
    if assigned(GetControl('RESPONSABLE')) then
      begin
      if GetControlText('RESPONSABLE') <> '' then
        begin
        ListeCombos := GetControlText('RESPONSABLE');
        if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
        ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
        ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
        StWhere:=StWhere + ' AND RAC_INTERVENANT in '+ListeCombos;
        end;
      end;
   end;
StWhere := StWhere + '))';
if stProduitpgi = 'GRC' then Confid:='CON' else Confid:='CONF';
StWhere := StWhere + RTXXWhereConfident(Confid);
SetControlText ('XX_WHERE',StWhere);
end;

procedure TOF_PROS_SSACTIONS.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  Case Key of
    VK_F10,VK_F9 :  if GetControlText ('TYPEACTION') = '' then key :=0;
  end;
if (ecran <> nil) then TFSTAT(ecran).FormKeyDown(Sender,Key,Shift);
end;

procedure TOF_PROS_SSACTIONS.TVOnDblClickCell(Sender: TObject );
begin
with TTobViewer(sender) do
  if stProduitpgi = 'GRF' then AGLLanceFiche('GC','GCFOURNISSEUR','',AsString[ColIndex('T_AUXILIAIRE'), CurrentRow],'MONOFICHE;T_NATUREAUXI=FOU')
  else AGLLanceFiche('GC','GCTIERS','',AsString[ColIndex('T_AUXILIAIRE'), CurrentRow],'MONOFICHE;T_NATUREAUXI='+AsString[ColIndex('T_NATUREAUXI'), CurrentRow])
end;

Initialization
registerclasses([TOF_PROS_SSACTIONS]);

end.
