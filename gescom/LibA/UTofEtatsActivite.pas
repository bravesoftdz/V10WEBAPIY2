unit UTofEtatsActivite;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,AGLInit,
      SaisUtil,HPanel,grids, Dicobtp, HSysMenu,HTB97,utofbaseetats,
{$IFDEF EAGLCLIENT}
   eMul,Maineagl,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, Mul, FE_Main,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF, FactUtil,
      ParamSoc,ActiviteUtil,utob;
Type
     TOF_ETATS_ACT = Class (TOF_BASE_ETATS)
      procedure OnArgument(stArgument : String ) ; override ;
      procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
//      procedure DateActiviteChange(Sender: TObject);
      procedure OnLoad; override ;
      procedure TableauObjectsInvisibles(Crit:string; var iNbChamps:integer; var tbChamps:PString);override ;
     public
      DateActivite    :  THEdit;
      DateActivite_   :  THEdit;
        procedure OnClose ; override ;

     END ;

const
NbTbChInv = 6 ;
TbChampsInvisibles : array[1..NbTbChInv] of string 	= (
          {1}        'ACT_FONCTIONRES',
          {2}        'TACT_FONCTIONRES',
          {3}        'ACT_FONCTIONRES_',
          {4}        'TACT_FONCTIONRES_',
          {5}        'ACT_TYPERESSOURCE',
          {6}        'TACT_TYPERESSOURCE'
          );

Procedure AFLanceFiche_Synthese_F_NF;
Procedure AFLanceFiche_Synthese_Clt_Art;
Procedure AFLanceFiche_Synthese_Clt_Ress;
Procedure AFLanceFiche_Synthese_Ress_Art;
Procedure AFLanceFiche_Synthese_Ress_Aff;

implementation

procedure TOF_ETATS_ACT.OnArgument(stArgument : String );
begin
// Gestion automatique de l'alimentation du champ DateActivite_ par rapport au DateActivite
DateActivite := THEdit(GetControl('ACT_DATEACTIVITE'));
DateActivite_ := THEdit(GetControl('ACT_DATEACTIVITE_'));

if (DateActivite<>nil) and (DateActivite_<>nil) then
   begin
//   DateActivite.OnChange:=DateActiviteChange;
   // Tenir compte de la valeur par defaut
//   if (DateActivite.Text<>'') and (DateActivite.Text<>' ') then
//      DateActiviteChange(DateActivite);
   end;
if (THEdit(GetControl('ACT_ETATVISA'))<>NIl) and (GetParamSoc('so_afVisaActivite') = False) then begin
   SetControlVisible('ACT_ETATVISA',False);
   SetControlVisible('TACT_ETATVISA',False);
   SetControlText ('ACT_ETATVISA','ATT;VIS');
   end;

Inherited;
setContrOlVisible ('BPROG',FALSE); // les état utilisent fct ConversionUnite.. pas OK plannification
End;

(*procedure TOF_ETATS_ACT.DateActiviteChange(Sender: TObject);
begin
DateActivite_.Text := DateToStr(PlusDate(StrToDate(DateActivite.Text),12,'M'));
end; *)

procedure TOF_ETATS_ACT.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Inherited;
Aff := Nil; Aff_ := Nil;
end;

procedure TOF_ETATS_ACT.OnClose;
BEGIN
if (TheTOB<>nil) then begin TheTOB.Free; TheTOB:=nil; end;
inherited;
END;

procedure TOF_ETATS_ACT.OnLoad;
var
EtatAct:THEdit;
begin
if (TheTOB<>nil) then begin TheTOB.Free; TheTOB:=nil; end;
if (ecran.name)<>'AFSYNTH_F_NF' then DateActivite_.Text := DateToStr(PlusDate(StrToDate(DateActivite.Text),12,'M'));
inherited;

EtatAct:=THEdit(GetControl('ACT_ETATVISA'));
if (EtatAct<>nil) then
if (EtatAct.visible=false) then
   SetControlText('ACT_ETATVISA','ATT;VIS');
end;

procedure TOF_ETATS_ACT.TableauObjectsInvisibles(Crit:string; var iNbChamps:integer; var tbChamps:PString);
begin
if (Crit='RESS') then
    begin
    iNbChamps := NbTbChInv;
    tbChamps := @TbChampsInvisibles;
    end;
end;

Procedure AFLanceFiche_Synthese_F_NF;
begin
AGLLanceFiche ('AFF','AFSYNTH_F_NF','','','');
end;
Procedure AFLanceFiche_Synthese_Clt_Art;
begin
AGLLanceFiche ('AFF','AFSYNTHCLIART','','','');
end;
Procedure AFLanceFiche_Synthese_Clt_Ress;
begin
AGLLanceFiche ('AFF','AFSYNTHCLIENT','','','');
end;
Procedure AFLanceFiche_Synthese_Ress_Aff;
begin
AGLLanceFiche ('AFF','AFSYNTHRESS','','','');
end;
Procedure AFLanceFiche_Synthese_Ress_Art;
begin
AGLLanceFiche ('AFF','AFSYNTHRESSART','','','');
end;


Initialization
registerclasses([TOF_ETATS_ACT]);
end.
