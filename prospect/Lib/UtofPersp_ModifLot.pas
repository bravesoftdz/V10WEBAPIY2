unit UTofPersp_ModifLot;

interface
uses UTOF,GCMZSUtil,HCtrls,
{$IFDEF EAGLCLIENT}
     eMul,MainEAGL,
{$ELSE}
     Mul,Fe_Main,
{$ENDIF}
{$IFDEF GIGI}
     ParamSoc,
{$ENDIF}
{$ifdef AFFAIRE}
UTOFAFTRADUCCHAMPLIBRE,
{$endif}
     Forms,Classes,M3FP,HMsgBox,Controls,HEnt1,UtilGc,EntRT,SysUtils,
     UtilRT;
Type
{$ifdef AFFAIRE}
                //PL le 18/05/07 pour faire affectation depuis ressource si paramétré
     TOF_Persp_ModifLot = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_Persp_ModifLot = Class (TOF)
{$endif}
       private
       procedure ModificationParLotDesPropositions;
       public
       Origine,TBlName : string;
       procedure OnLoad ; override ;
       procedure OnClose ; override ;
       procedure OnArgument (stArgument : string); override;
     end;

procedure AGLModifParLotDesPropositions(parms : array of variant; nb : integer);
Procedure RTLanceFiche_Persp_ModifLot(Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation
uses UtilSelection
{CRM_MNG_012FQ10807_070408 }
,wcommuns,StdCtrls ;


Procedure RTLanceFiche_Persp_ModifLot(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
  AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Persp_ModifLot.OnArgument (stArgument : string);
var arg,Ctrl : String;
begin
inherited;
arg:=stArgument;
Origine:=ReadTokenSt(arg);
Ctrl := 'Z';

if (Origine = 'RPE') then
begin
  MulCreerPagesCL(Ecran,'NOMFIC=RTPERSPECTIVES');
  Ecran.Caption := TraduireMemoire('Modification en série des infos complémentaires propositions');
  UpdateCaption(Ecran);
end;

{ je remet à blanc car fait planter le paramétrage liste }
V_PGI.ExtendedFieldSelection:='';

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
{$Ifdef GIGI}
 SetControlVisible('RPE_REPRESENTANT',false);
 SetControlVisible('TRPE_REPRESENTANT',false);
 if (GetControl('RPE_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    end;
{$endif}
end;

procedure TOF_Persp_ModifLot.OnLoad;
var xx_where : string;
{CRM_MNG_012FQ10807_070408 }
 wheremulti : string;

begin
  xx_where := '';
  //if (TCheckbox(TFMUL (Ecran).FindComponent('PROPOPRINCIPALE')).Checked = true) then
  if GetControlText('PROPOPRINCIPALE') = 'X' then
     xx_where := '(RPE_VARIANTE=0 or RPE_PERSPECTIVE=RPE_VARIANTE)';
  if (VH_RT.RTCreatPropositions = False) and (Trim(RTXXWhereConfident('MOD')) <> '') then
    xx_where := xx_where + RTXXWhereConfident('MOD')
  else
    xx_where := xx_where + RTXXWhereConfident('CON');

  TBlName := 'PERSPECTIVES';
  if (Origine = 'RPE') then
    TBlName := 'RTINFOS00V';

  SetControlText('XX_WHERE',xx_where) ;
{CRM_MNG_012FQ10807_070408 }
  wheremulti:=MulWhereMultiChoix (TForm (Ecran),'RDV',iif( Assigned(TRadioButton(GetControl('MULTIET'))) AND
    (TRadioButton(GetControl('MULTIET')).checked), 'AND','OR'));
  if Assigned(GetControl('XX_WHEREMULTI')) then
    SetControlText('XX_WHEREMULTI',wheremulti);
{fin CRM_MNG_012FQ10807_070408 }
  
end;

procedure TOF_Persp_ModifLot.OnClose;
begin
V_PGI.ExtendedFieldSelection:='' ;
end;

/////////////// ModificationParLotDesPropositions //////////////
procedure TOF_Persp_ModifLot.ModificationParLotDesPropositions;
Var F : TFMul ;
    Parametrages : String;
    TheModifLot : TO_ModifParLot;
begin
{ mng : je remets ce bloc ici car "Paramétrer la liste" remet à blanc V_PGI.ExtendedFieldSelection }

V_PGI.ExtendedFieldSelection:='Z';

F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin
  V_PGI.ExtendedFieldSelection:='' ;
  MessageAlerte('Aucun élément sélectionné');
  exit;
  end;
{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
   if not F.Fetchlestous then
     begin
     F.bSelectAllClick(Nil);
     F.bSelectAll.Down := False;
     V_PGI.ExtendedFieldSelection:='' ;
     exit;
     end else
     begin
     F.bSelectAll.Down := True;
     F.Fliste.AllSelected := true;
     end;
{$ENDIF}

TheModifLot := TO_ModifParLot.Create;
TheModifLot.F := F.FListe; TheModifLot.Q := F.Q;
TheModifLot.NatureTiers := Origine;
TheModifLot.Titre := Ecran.Caption;
TheModifLot.stAbrege :='';
TheModifLot.TableName:=TBlName;
TheModifLot.FCode := 'RPE_PERSPECTIVE';
TheModifLot.Nature := 'RT';
TheModifLot.FicheAOuvrir := 'RTPERSPECTIVES';
TheModifLot.StContexte := Origine;

if Origine = 'RPE' then
  TheModifLot.FicheAOuvrir := 'RTPARAMCL';

ModifieEnSerie(TheModifLot, Parametrages) ;

if F.bSelectAll.Down then
    begin
    F.bSelectAllClick(Nil);
    F.bSelectAll.Down := False;
    end;
V_PGI.ExtendedFieldSelection:='' ;
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLModifParLotDesPropositions(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_Persp_ModifLot) then TOF_Persp_ModifLot(TOTOF).ModificationParLotDesPropositions else exit;
end;

Initialization
registerclasses([TOF_Persp_ModifLot]) ;
RegisterAglProc('ModifParLotDesPropositions',TRUE,0,AGLModifParLotDesPropositions);
end.
