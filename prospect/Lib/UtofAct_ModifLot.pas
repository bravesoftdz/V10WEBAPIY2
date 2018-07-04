unit UTofAct_ModifLot;

interface
uses UTOF,GCMZSUtil,HCtrls,
{$IFDEF EAGLCLIENT}
     eMul,MainEAGL,
{$ELSE}
     Mul,Fe_Main,
{$ENDIF}
     ParamSoc,
     Forms,Classes,M3FP,HMsgBox,Controls,HEnt1,EntRT,StdCtrls,
     UtilRT;
Type
     TOF_Act_ModifLot = Class(TOF)
       private
       procedure ModificationParLotDesActions;
       public
       Origine : string;
       TBlName : String;      //nom de la table pour Maj
       procedure OnLoad ; override ;
       procedure OnClose ; override ;
       procedure OnArgument (stArgument : string); override;
     end;

procedure AGLModifParLotDesActions(parms : array of variant; nb : integer);
Procedure RTLanceFiche_Act_ModifLot(Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation
{CRM_MNG_012FQ10807_070408 }
uses wcommuns,utilselection ;

Procedure RTLanceFiche_Act_ModifLot(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
  AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Act_ModifLot.OnArgument (stArgument : string);
begin
inherited;
Origine:=stArgument;
{ je remet à blanc car fait planter le paramétrage liste }
V_PGI.ExtendedFieldSelection:='';

if Origine = 'GRCCOMPL' then
begin
  { CRM_20080407_MNG_012;10372_DEB }
//  MulCreerPagesCL(Ecran,'NOMFIC=RTACTIONS');                  //FQ 10815    //TJA 11/09/2008
  Ecran.Caption :=  'Modification en série des infos complémentaires actions';
  UpdateCaption(Ecran);
end;
  if GetParamSocSecur('SO_RTGESTINFOS001',False) = True then   //FQ 10815     //TJA 11/09/2008
    MulCreerPagesCL(Ecran,'NOMFIC=RTACTIONS');

{$Ifdef GIGI}
 if (GetControl('RAC_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RAC_OPERATION',false);
    SetControlVisible('TRAC_OPERATION',false);
    end;
{$endif}
SetControlProperty('RAC_ETATACTION', 'PLUS', THDBValComboBox(GetControl('RAC_ETATACTION')).Plus+' AND CO_CODE <> "ZCL"');
end;

procedure TOF_Act_ModifLot.OnLoad;
var stWhere : string;
{CRM_MNG_012FQ10807_070408 }
    wheremulti : string;
begin
if TCheckBox(GetControl('AVECCHAINAGE')).state=cbgrayed then
   stwhere:=''
 else
   if TCheckBox(GetControl('AVECCHAINAGE')).checked=False then stwhere:='RAC_NUMCHAINAGE = 0'
   else  stwhere:='RAC_NUMCHAINAGE <> 0';

TBlName := 'ACTIONS';
if (Origine = 'GRC') then stWhere := stWhere + RTXXWhereConfident('CON')
else if Origine = 'GRCCOMPL' then
begin
  stWhere := stWhere + RTXXWhereConfident('CON');
  TBlName := 'RTINFOS001';
end
else stWhere := stWhere + RTXXWhereConfident('CONF') ;
setControlText('XX_WHERE',stWhere);
{CRM_MNG_012FQ10807_070408 }
  wheremulti:=MulWhereMultiChoix (TForm (Ecran),'RD1',iif( Assigned(TRadioButton(GetControl('MULTIET'))) AND
    (TRadioButton(GetControl('MULTIET')).checked), 'AND','OR'));
  if Assigned(GetControl('XX_WHEREMULTI')) then
    SetControlText('XX_WHEREMULTI',wheremulti);
{fin CRM_MNG_012FQ10807_070408 }
end;

procedure TOF_Act_ModifLot.OnClose;
begin
V_PGI.ExtendedFieldSelection:='' ;
end;

/////////////// ModificationParLotDesPropositions //////////////
procedure TOF_Act_ModifLot.ModificationParLotDesActions;
Var F : TFMul ;
    Parametrages : String;
    TheModifLot : TO_ModifParLot;
begin
{ mng : je remets ce bloc ici car "Paramétrer la liste" remet à blanc V_PGI.ExtendedFieldSelection }

if Origine = 'GRCCOMPL' then V_PGI.ExtendedFieldSelection := 'Z'    // pour modif en lot des actions compl
else if (TFMul(Ecran).name = 'RTACT_MODIFLOT') then V_PGI.ExtendedFieldSelection:='2'
else V_PGI.ExtendedFieldSelection:='1';

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
TheModifLot.NatureTiers := 'ACT';
TheModifLot.Titre := Ecran.Caption;
TheModifLot.stAbrege :='';
TheModifLot.TableName:= TBlName;
TheModifLot.FCode := 'RAC_AUXILIAIRE;RAC_NUMACTION';
TheModifLot.Nature := 'RT';
TheModifLot.FicheAOuvrir := 'RTACTIONS';
TheModifLot.StContexte := Origine;

{ CRM_20080407_MNG_012;10372_DEB }
if Origine = 'GRCCOMPL' then
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
procedure AGLModifParLotDesActions(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_Act_ModifLot) then TOF_Act_ModifLot(TOTOF).ModificationParLotDesActions else exit;
end;

Initialization
registerclasses([TOF_Act_ModifLot]) ;
RegisterAglProc('ModifParLotDesActions',TRUE,0,AGLModifParLotDesActions);
end.
