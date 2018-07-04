unit UTofAfActiviteCube;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,paramsoc,
{$IFDEF EAGLCLIENT}
  Maineagl,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, FE_Main,
{$ENDIF}
  HCtrls,UTOF,UTofAfBaseCodeAffaire,AfUtilArticle,Dicobtp ,UtilGc
//  , UtofAftableaubord
  ;

Type

TOF_AfActiviteCube = Class (TOF_AFBASECODEAFFAIRE)
  procedure OnArgument (stArgument : string); override;
  procedure NomsChampsAffaire(var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);override ;
  END;

Procedure AFLanceFiche_ActiviteCube;

implementation

//******************************************************************************
//**************** Cube d'activité  ********************************************
//******************************************************************************

Procedure AFLanceFiche_ActiviteCube;
begin
  AGLLanceFiche ('AFF','AFACTIVITECUBE','','','');
end;

procedure TOF_AfActiviteCube.OnArgument(stArgument : String );
Var ComboTypeArticle : THMultiValComboBox;
Zone: THValComboBox;
ComboType : THValComboBox;
Critere : string;
Begin
Inherited;
 // mcd 10/12/02 pour limiter zone fournisseur
   // mcd 20/01/03 suppression du multival. attention. ne pas remettre sinon impossible
   // de faire une sélection sur les champs vide (autre que FF pe pour GI). voir sinon comment on peut faire ... passer par une tablette spécifique ??? avec vide plus les valeurs ???
   //Zone :=ThMultiValCOmboBox(GetControl('ACT_NATUREPIECEG'));
  Zone :=ThValCOmboBox(GetControl('ACT_NATUREPIECEG'));
  If zone <>Nil then Zone.plus := Zone.plus +AfPlusNatureAchat;
  if (GetParamSoc('so_afVisaActivite') = False) then
  begin
     SetControlVisible('ACT_ETATVISA',False);
     SetControlVisible('TACT_ETATVISA',False);
  end;
  if (GetParamSoc('SO_AFAPPPOINT') = False) then
  begin
     SetControlVisible('ACT_ETATVISAFAC',False);
     SetControlVisible('TACT_ETATVISAFAC',False);
  end;
  ComboTypeArticle:=THMultiValComboBox(GetControl('ACT_TYPEARTICLE'));
  ComboTypeArticle.plus:=PlusTypeArticle;
  if ComboTypeArticle.Text='' then ComboTypeArticle.Text:=PlusTypeArticleText;
  if Not(GetParamSOc('SO_AFAPPAVECBM')) then begin
      ComboType:=THValComboBox(GetControl('ACT_TYPEACTIVITE'));
      ComboType.plus:=' AND CO_CODE<>"BON"';
      ComboType.value:='REA';
      end;
   If Not GetParamSoc ('So_AfAppPoint') then
     begin
     SetControlVisible ('ACT_ETATVISAFAC',False);
     SetControlVisible ('TACT_ETATVISAFAC',False);
     end;
  Ecran.Caption := TraduitGA('Cube activité');
  UpdateCaption(Ecran);
End;


procedure TOF_AfActiviteCube.NomsChampsAffaire(var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('ACT_AFFAIRE'));
  Aff1:=THEdit(GetControl('ACT_AFFAIRE1'));
  Aff2:=THEdit(GetControl('ACT_AFFAIRE2'));
  Aff3:=THEdit(GetControl('ACT_AFFAIRE3'));
  Aff4:=THEdit(GetControl('ACT_AVENANT'));
  Tiers:=THEdit(GetControl('ACT_TIERS'));
end;

Initialization
registerclasses([TOF_AfActiviteCube]);
end.
