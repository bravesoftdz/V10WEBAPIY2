unit UTofPerspective;

interface
uses  StdCtrls,Classes,forms,sysutils,
      HCtrls,HEnt1,UTOF,HMsgBox
      ,UtilRT ,UtilGC, EntRT, HTB97
{$IFDEF EAGLCLIENT}
     ,eMul,MainEagl
{$ELSE}
     ,Mul,Fe_Main
{$ENDIF}
{$ifdef AFFAIRE}
,UTOFAFTRADUCCHAMPLIBRE
{$endif}
     ,AGLInit
;
Type
{$ifdef AFFAIRE}
                //PL le 18/05/07 pour faire affectation depuis ressource si paramétré
     TOF_PERSP_MUL = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_PERSP_MUL = Class (TOF)
{$endif}
     private
         //procedure RecupWhere;
        stArguments : String;
        procedure FLISTE_OnDblClick(Sender: TObject);
    		procedure Binsert_OnClick(Sender: TObject);
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override;
     END ;
var stWhere: string ;
    VerrouModif :boolean;
implementation
uses
  UtilSelection,ParamSoc, uTOFComm;

procedure TOF_PERSP_MUL.OnArgument(Arguments : String ) ;
var i:integer;
begin
inherited ;
  if (ecran <> Nil) and GetParamSocSecur('SO_RTGESTINFOS00V',False) then
    MulCreerPagesCL(Ecran,'NOMFIC=RTPERSPECTIVES');

  VerrouModif:=False;
  stArguments:=Arguments;
  if not VH_RT.RTCreatPropositions then
    begin
    i := pos('CONSULTATION',Arguments);
    if i <> 0 then VerrouModif :=true;
    end;
for i:=1 to 3 do
    SetControlCaption('TRPE_TABLELIBREPER'+IntToStr(i),'&'+RechDom('RTLIBCHAMPSLIBRES','PL'+IntToStr(i),FALSE)) ;

SetControlText ('XX_WHERE', 'RPE_VARIANTE=0 or RPE_PERSPECTIVE=RPE_VARIANTE');
if (TFMUL (Ecran).name = 'RTPERSPECTIVE_REC') then
    begin
    i:=pos('=',Arguments);
    if i<>0 then     //Montant:=StrToFloat(copy(S,x+1,length(S)));
        SetControlText ('MONTANTPROPAL',copy(Arguments,i+1,length(Arguments)));
    end;
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
    end;
  end;
  if Assigned(GetControl('FLISTE')) then
    THGrid(GetControl('FLISTE' )).OnDblClick := FLISTE_OnDblClick;

{$ifdef GIGI}
 SetControlVisible('RPE_REPRESENTANT',false);
 SetControlVisible('TRPE_REPRESENTANT',false);
 If (Not GetParamSocSecur ('SO_RTPROJGESTION',False)) then
   begin
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
   end;
 if (GetControl('RPE_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    end;
{$ENDIF}

{$IFDEF GRCLIGHT}
  if not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
    begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
    end;
{$ENDIF GRCLIGHT}
    TToolBarButton97(getCOntrol('Binsert')).OnClick  := Binsert_OnClick;
end;

procedure TOF_PERSP_MUL.FLISTE_OnDblClick(Sender: TObject);
var stArg : String;
begin
{$ifdef AFFAIRE}
//  Inherited ; // PL le 19/05/07 : doit surcharger l'événement onDblClick de la TofComm
// et non pas le dériver...
 {$else}
  Inherited ;
{$endif}
  if GetField('RPE_TIERS') = '' then exit;
  if (TFMul(Ecran).Name<>'RTPERSPECTIVE_REC') and (TFMul(Ecran).Name<>'RTPERSP_AFF') then
    begin
    stArg:='ACTION=MODIFICATION';
    if (RTDroitModiftiers(GetField('RPE_TIERS'))=False)  and ( RTXXWhereConfident('CREATP') = '-' )
    then stArg:= 'ACTION=CONSULTATION';
    if Assigned(GetControl('BVERROU')) and ( TCheckBox(GetControl('BVERROU')).Checked ) then
      stArg:= 'ACTION=CONSULTATION';


    if pos('FICHEPROPOSITION',stArguments) <> 0 then stArg:=stArg+';FICHEPROPOSITION';
    if pos('FICHEACTION',stArguments) <> 0 then stArg:=stArg+';FICHEACTION';

    if (TFMul(Ecran).Name='RTPERSP_MUL_TIERS') then
      begin
      if GetControlText('RPE_NUMEROCONTACT') <> '' then
           stArg:=stArg+';FICHECONTACT'+';RPE_NUMEROCONTACT='+GetControlText('RPE_NUMEROCONTACT')
         else stArg:=stArg+';FICHETIERS';
      stArg:=stArg+';NOCHANGEPROSPECT;RPE_TIERS='+GetControlText('RPE_TIERS');
      end
    else
      if pos('FICHETIERS',stArguments) <> 0 then stArg:=stArg+';FICHETIERS';
    AglLancefiche('RT','RTPERSPECTIVES','',GetField('RPE_PERSPECTIVE'),stArg);
    AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
    end
  else
    begin
    TFMul(Ecran).Retour := IntToStr(GetField('RPE_PERSPECTIVE'));
    if (TFMul(Ecran).Name='RTPERSPECTIVE_REC') then
       TFMul(Ecran).Retour := TFMul(Ecran).Retour +';1';
    TFMul(Ecran).Close;
    end;
end;


procedure TOF_PERSP_MUL.OnLoad;
var xx_where : string;
begin
Inherited;
xx_where := '';
if (TCheckbox(TFMUL (Ecran).FindComponent('PROPOPRINCIPALE')).Checked = true) then
   xx_where := '(RPE_VARIANTE=0 or RPE_PERSPECTIVE=RPE_VARIANTE)';

if (TFMUL (Ecran).name = 'RTPERSPECTIVE_MUL') then xx_where := xx_where + RTXXWhereConfident('CON') else
begin
  SetControlEnabled('BINSERT',not VerrouModif);
  SetControlChecked('BVERROU',VerrouModif);
end;

SetControlText('XX_WHERE',xx_where) ;
end;

procedure form_where;
begin
while pos (#$D, stWhere) > 0 do delete (stWhere, pos (#$D, stWhere), 1);
while pos (#$A, stWhere) > 0 do delete (stWhere, pos (#$A, stWhere), 1);
end;

procedure TOF_PERSP_MUL.Binsert_OnClick(Sender : TObject);
var stArg : string;
begin
  if AGLRTExisteConfident([],0) <> '' then
  begin
    stArg:='ACTION=CREATION';

    if (THEdit(GetCOntrol('RPE_TIERS')).Text<>'') then stArg:=stArg+';RPE_TIERS='+THEdit(GetCOntrol('RPE_TIERS')).text;
    if (THEdit(GetCOntrol('RPE_PROJET')).Text<>'') then stArg:=stArg+';RPE_PROJET='+THEdit(GetCOntrol('RPE_PROJET')).text;
    if (THEdit(GetCOntrol('RPE_NUMEROACTION')).Text<>'') then stArg:=stArg+';RPE_NUMEROACTION='+THEdit(GetCOntrol('RPE_NUMEROACTION')).text;
    if pos('NOCHANGEPROSPECT',stargument)>0 then stArg:=stArg+';NOCHANGEPROSPECT' ;

    if pos('FICHETIERS',StArgument)>0 then stArg:=stArg+';FICHETIERS';
    if pos('FICHEPROPOSITION',StArgument)>0 then stArg:=stArg+';FICHEPROPOSITION';
    if pos('FICHEACTION',StArgument)>0 then stArg:=stArg+';FICHEACTION';

    AglLanceFiche('RT','RTPERSPECTIVES','','',stArg) ;
  	TtoolBarButton97(GetCOntrol('Bcherche')).Click;
  end
  else
  	PGiBox ('Vous n''avez pas les droits d''accès en création');
end;

{mng pour eagl procedure TOF_PERSP_MUL.RecupWhere;
var TSql : TQuery;
    p : integer;
begin
stwhere:='';
TSql := TQuery.Create (nil);
TSql.SQL.Add ('');
RecupWhereSQL (TFMUL(Ecran).Q, TSql);
stWhere := TSql.Text;
form_where;
TSql.Free;
end;
}

Initialization
registerclasses([TOF_PERSP_MUL]);
end.

