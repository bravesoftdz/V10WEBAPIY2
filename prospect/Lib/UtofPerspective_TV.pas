unit UtofPerspective_TV;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,
      HCtrls,HEnt1,UTOF, Stat ,UtilGC
      , UTobView,  UtilSelection ,UtilRT ,EntRT, EntGC,TiersUtil
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
                //mcd 11/05/2006 12940  pour faire affectation depuis ressource si paramétré
     TOF_PERSP_TV = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_PERSP_TV = Class (TOF)
{$endif}
     private
         TobViewer1: TTobViewer;
         procedure TVOnDblClickCell(Sender: TObject ) ;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
     END ;

implementation
uses ParamSoc;

procedure TOF_PERSP_TV.OnArgument(Arguments : String ) ;
var F : TForm;
    Memo : TMemo;
begin
inherited ;
  F := TForm (Ecran);
  MulCreerPagesCL(F,'NOMFIC=GCTIERS');
  if (ecran <> Nil) and GetParamSocSecur('SO_RTGESTINFOS00V',False) then
    MulCreerPagesCL(Ecran,'NOMFIC=RTPERSPECTIVES');

TobViewer1:=TTobViewer(getcontrol('TV'));
TobViewer1.OnDblClick:= TVOnDblClickCell ;
if (F.name = 'RTPERSPSSDEV_TV') then
   begin
   THEdit(F.FindComponent('RESPONSABLE')).Text := VH_RT.RTResponsable;
{$IFDEF AFFAIRE}
   if (( not (ctxAffaire in V_PGI.PGIContexte) ) and ( not ( ctxGCAFF in V_PGI.PGIContexte) ) ) or
   ( VH_GC.GASeria=false ) then
{$ENDIF}   
      begin
      SetControlVisible ('SANSAFFAIRE',false);
      //Ecran.Caption:='Propositions / Devis            ' ;
      TFStat(Ecran).Caption := TraduireMemoire('Propositions / Devis            ');
      UpdateCaption(Ecran);
      end;
   end;
memo := TMemo(GetControl('FSQL'));
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
   else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    if not (ctxscot in V_PGI.PGICOntexte) then
       begin
       SetControlVisible ('T_MOISCLOTURE',false);
       SetControlVisible ('T_MOISCLOTURE_',false);
       SetControlVisible ('TT_MOISCLOTURE',false);
       SetControlVisible ('TT_MOISCLOTURE_',false);
       end;
    If (Ecran).name ='RTPERSPECTIVE_TV' then Memo.lines[4]:=',YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3'
       else If (Ecran).name ='RTPERSPSSDEV_TV' then Memo.lines[4]:=',YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3';
    end;
  end;
{$Ifdef GIGI}
 if (GetControl('RPE_REPRESENTANT') <> nil) then  SetControlVisible('RPE_REPRESENTANT',false);
 if (GetControl('TRPE_REPRESENTANT') <> nil) then  SetControlVisible('TRPE_REPRESENTANT',false);
 if (GetControl('RPE_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    end;
 if (GetControl('RPE_PROJET') <> nil) and Not GetParamsocSecur('SO_RTPROJGESTION',False)
    then  begin
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
    end;
 if (GetControl('RPE_TYPETIERS') <> nil)
    then  begin
    SetControlText ('TRPE_TYPETIERS', TraduireMemoire('Nature tiers'));
    SetControlProperty ('RPE_TYPETIERS', 'Complete', true);
    SetControlProperty ('RPE_TYPETIERS', 'Datatype', 'TTNATTIERS');
    SetControlProperty ('RPE_TYPETIERS', 'Plus', VH_GC.AfNatTiersGRCGI);
    end;
 If (Ecran).name ='RTPERSPECTIVE_TV' then
    begin
    Memo.lines[5]:=',ARS_LIBELLE as Responsable';
    Memo.lines[10]:=',T_PRESCRIPTEUR,T_SOCIETEGROUPE,T_APE,T_MOISCLOTURE';
    Memo.lines[32]:='';
    end
    else If (Ecran).name ='RTPERSPCON_TV' then
      begin
      Memo.lines[5]:='';
      Memo.lines[6]:='ARS_LIBELLE as Responsable,';
      Memo.lines[10]:='T1.T_CODEPOSTAL';  //suppression de tout ce qui concerne Tuerscompl.. non saisi en GI dans concurrent
      Memo.lines[11]:='';
      Memo.lines[12]:='';
      Memo.lines[13]:='';
      Memo.lines[31]:='';
      Memo.lines[33]:='';
      end
    else If (Ecran).name ='RTPERSPSSDEV_TV' then
      begin
      Ecran.Caption := TraduireMemoire('Propositions/proposition de missions ');
      UpdateCaption(Ecran);
      Memo.lines[5]:=',ARS_LIBELLE as Responsable,T_MOISCLOTURE';
      Memo.lines[38]:='';
      end;
{$endif}
{$IFDEF GRCLIGHT}
  if not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
    begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
    end;
{$ENDIF GRCLIGHT}

end;

procedure TOF_PERSP_TV.OnLoad;
Var F : TForm;
    xx_where,DebdatePiece,FindatePiece,stSql : string  ;
    Deb_Piece,Fin_Piece : TDateTime;
begin
inherited;
F := TForm (Ecran);
if (TCheckbox(F.FindComponent('PROPOPRINCIPALE')).Checked = true) then
   xx_where := '(RPE_VARIANTE=0 or RPE_PERSPECTIVE=RPE_VARIANTE)'
else
   xx_where := '';
{ Analyse concurrence : pas de confidentialité car 2 left join tiers T1 T2 et pas ocmpatible }
if (F.name <> 'RTPERSPCON_TV')  then
  xx_where := xx_where + RTXXWhereConfident('CON');

if (F.name = 'RTPERSPSSDEV_TV')  then
begin
  DebdatePiece := GetControlText('DATEPIECE');
  if not (IsValidDate(DebdatePiece)) then
  begin
    Deb_Piece := DebutDeMois (V_PGI.DateEntree);
    SetControlText( 'DATEPIECE' , DateToStr (Deb_Piece));
  end
  else Deb_Piece := StrToDate (DebdatePiece);
  FindatePiece := GetControlText('DATEPIECE_');
  if not (IsValidDate(FindatePiece)) then
  begin
    Fin_Piece := FinDeMois (V_PGI.DateEntree);
    SetControlText( 'DATEPIECE_' , DateToStr (Fin_Piece));
  end
  else Fin_Piece := StrToDate (FindatePiece);

  if (TCheckbox(F.FindComponent('SANSDEVIS')).state=cbUnChecked) then
  xx_where := xx_where + ' AND (RPE_PERSPECTIVE not in '+
//  '(SELECT GP_PERSPECTIVE from PIECE where GP_NATUREPIECEG="DE" and GP_VENTEACHAT="VEN" and GP_PERSPECTIVE<>"" '+
  '(SELECT GP_PERSPECTIVE from PIECE where GP_NATUREPIECEG="DE" and GP_VENTEACHAT="VEN" and GP_PERSPECTIVE=RPE_PERSPECTIVE '+
  ' AND GP_DATEPIECE >= "'+USDateTime (Deb_Piece)+'" AND GP_DATEPIECE <= "'+USDateTime (Fin_Piece)+ '" ))'
  else if (TCheckbox(F.FindComponent('SANSDEVIS')).state=cbChecked) then
  xx_where := xx_where + ' AND (RPE_PERSPECTIVE in '+
//  '(SELECT GP_PERSPECTIVE from PIECE where GP_NATUREPIECEG="DE" and GP_VENTEACHAT="VEN" and GP_PERSPECTIVE<>"" '+
  '(SELECT GP_PERSPECTIVE from PIECE where GP_NATUREPIECEG="DE" and GP_VENTEACHAT="VEN" and GP_PERSPECTIVE=RPE_PERSPECTIVE '+
  ' AND GP_DATEPIECE >= "'+USDateTime (Deb_Piece)+'" AND GP_DATEPIECE <= "'+USDateTime (Fin_Piece)+ '" ))';

  if ( (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) ) and
  ( VH_GC.GASeria=True ) then
    begin
    if (TCheckbox(F.FindComponent('SANSAFFAIRE')).state=cbUnChecked) then
    xx_where := xx_where + ' AND (RPE_PERSPECTIVE not in '+
    '(SELECT AFF_PERSPECTIVE from AFFAIRE where AFF_PERSPECTIVE=RPE_PERSPECTIVE ))'
    else if (TCheckbox(F.FindComponent('SANSAFFAIRE')).state=cbChecked) then
    xx_where := xx_where + ' AND (RPE_PERSPECTIVE in '+
    '(SELECT AFF_PERSPECTIVE from AFFAIRE where AFF_PERSPECTIVE=RPE_PERSPECTIVE ))';
    end;
end;

if (F.name = 'RTPERSPECTIVE_TV')  then
begin
  stSql := '(SELECT RPT_PERSPECTIVE from PERSPECTIVESTIERS where RPT_PERSPECTIVE=RPE_PERSPECTIVE ';
  if (TCheckbox(F.FindComponent('SANSCONCURRENT')).state<>cbGrayed) and (GetControlText('CONCURRENT') <> '') then
  begin
  stSql := stSql + ' AND RPT_TIERS = "'+GetControlText('CONCURRENT')+'"';
  end;
  if (TCheckbox(F.FindComponent('SANSCONCURRENT')).state=cbUnChecked) then
  xx_where := xx_where + ' AND (RPE_PERSPECTIVE not in '+ stSql + ' ))'
  else if (TCheckbox(F.FindComponent('SANSCONCURRENT')).state=cbChecked) then
  xx_where := xx_where + ' AND (RPE_PERSPECTIVE in '+ stSql + ' ))';
end;
SetControlText('XX_WHERE',xx_where) ;
end;

procedure TOF_PERSP_TV.TVOnDblClickCell(Sender: TObject );
var staction : string;
    F : TForm;
begin
F := TForm (Ecran);
with TTobViewer(sender) do
  begin
  staction:='ACTION=CONSULTATION';
  if RTDroitModifTiers(AsString[ColIndex('RPE_TIERS'), CurrentRow]) then staction:='ACTION=MODIFICATION';
  if ((ColName[CurrentCol] = 'RPE_AUXILIAIRE') or (ColName[CurrentCol] = 'RPE_TIERS')
    or (ColName[CurrentCol] = 'T_LIBELLE') or (ColName[CurrentCol] = 'RaisonSociale'))
{ Analyse concurrence : pas de confidentialité donc pas de zoom tiers }
    and (F.Name <> 'RTPERSPCON_TV') then
       V_PGI.DispatchTT (28,taConsult ,AsString[ColIndex('RPE_AUXILIAIRE'), CurrentRow], '','')
  else if (ColName[CurrentCol] = 'RPE_NUMEROACTION') and  (AsInteger[ColIndex('RPE_NUMEROACTION'), CurrentRow] <> 0 )
  then V_PGI.DispatchTT (22,taConsult ,AsString[ColIndex('RPE_AUXILIAIRE'), CurrentRow]+';'+IntToStr(AsInteger[ColIndex('RPE_NUMEROACTION'), CurrentRow]), '','')
  else if (ColName[CurrentCol] = 'RPE_OPERATION') and (trim(AsString[ColIndex('RPE_OPERATION'), CurrentRow])<>'')
  then  V_PGI.DispatchTT (23,taConsult ,AsString[ColIndex('RPE_OPERATION'), CurrentRow], '','')
  else if (ColName[CurrentCol] = 'RPE_PROJET') and (trim(AsString[ColIndex('RPE_PROJET'), CurrentRow])<>'')
  then V_PGI.DispatchTT (30,taConsult ,AsString[ColIndex('RPE_PROJET'), CurrentRow], '','')
  else if (ColName[CurrentCol] = 'T_SOCIETEGROUPE')  then  V_PGI.DispatchTT (8,taConsult ,AsString[ColIndex('T_SOCIETEGROUPE'), CurrentRow], '','')
  else if (ColName[CurrentCol] = 'T_PRESCRIPTEUR')  then  V_PGI.DispatchTT (8,taConsult ,AsString[ColIndex('T_PRESCRIPTEUR'), CurrentRow], '','')
  else if (ColName[CurrentCol] = 'CONCURRENT') or (ColName[CurrentCol] = 'NOM_CONCURRENT')
  then  V_PGI.DispatchTT (47,taConsult ,TiersAuxiliaire(AsString[ColIndex('Concurrent'), CurrentRow],False), '','T_NATUREAUXI=CON')
  else
   AGLLanceFiche('RT','RTPERSPECTIVES','',IntToStr(AsInteger[ColIndex('RPE_PERSPECTIVE'), CurrentRow]) ,staction);
  end;
end;

Initialization
registerclasses([TOF_PERSP_TV]);

end.
