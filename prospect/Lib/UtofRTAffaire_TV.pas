unit UtofRTAffaire_TV;

interface
uses  StdCtrls,Classes,forms,sysutils,
      HCtrls,HEnt1,UTOF,UtilGC
      , UTobView, UtilSelection ,UtilRT 
      ,COnfidentAffaire,HTB97
{$ifdef AFFAIRE}
      ,UtofAfTraducChampLibre
{$ELSE}
			,UtofAfBaseCodeAffaire
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
     TOF_RTAFFAIRE_TV = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_RTAFFAIRE_TV = Class (TOF_AFBASECODEAFFAIRE)
{$endif}
     procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override;
     procedure bSelectAff1Click(Sender: TObject);     override ;
     private
         TobViewer1: TTobViewer;
         procedure TVOnDblClickCell(Sender: TObject ) ;
         procedure Select2Click(Sender: TObject);
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
     END ;

implementation
uses
  ParamSoc ;

procedure TOF_RTAFFAIRE_TV.OnArgument(Arguments : String ) ;
var F : TForm;
Select2: TToolBarButton97;
begin
inherited ;
  F := TForm (Ecran);
  MulCreerPagesCL(F,'NOMFIC=GCTIERS');
  if GetParamSocSecur('SO_RTGESTINFOS00V',False) then
    MulCreerPagesCL(Ecran,'NOMFIC=RTPERSPECTIVES');

      TobViewer1:=TTobViewer(getcontrol('TV'));
  TobViewer1.OnDblClick:= TVOnDblClickCell ;
  Select2:=TToolBarButton97(getcontrol('BSELECTAFF1_'));
  Select2.OnClick:= Select2Click ;
  If Not (Ctxscot in V_PGI.PGICOntexte) then
    begin // MCd 08/02/2005 zones uniquement en GI !!
    SetControlVisible ('TT_MOISCLOTURE', False);
    SetControlVisible ('T_MOISCLOTURE', False);
    SetControlVisible ('TT_MOISCLOTURE_', False);
    SetControlVisible ('T_MOISCLOTURE_', False);
    SetControlVisible ('AFF_DATEDEBEXER', False);
    SetControlVisible ('TAFF_DATEDEBEXER', False);
    SetControlVisible ('AFF_DATEDEBEXER_', False);
    SetControlVisible ('TAFF_DATEDEBEXER_', False);
    SetControlVisible ('TAFF_DATEFINEXER', False);
    SetControlVisible ('AFF_DATEFINEXER', False);
    SetControlVisible ('TAFF_DATEFINEXER_', False);
    SetControlVisible ('AFF_DATEFINEXER_', False);
    end;
  if (GetControl('YTC_RESSOURCE1') <> nil)  then
    begin
    if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
    else GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    end;
end;

procedure TOF_RTAFFAIRE_TV.OnLoad;
Var F : TForm;
    xx_where: string  ;
begin
inherited;
F := TForm (Ecran);
xx_where:=GetControlText('XX_WHERE');
if (TCheckbox(F.FindComponent('PROPOPRINCIPALE')).Checked = true) then
   xx_where := xx_where+' AND (RPE_VARIANTE="0" or RPE_PERSPECTIVE=RPE_VARIANTE)';
xx_where := xx_where + RTXXWhereConfident('CON');

SetControlText('XX_WHERE',xx_where) ;
end;

procedure TOF_RTAFFAIRE_TV.TVOnDblClickCell(Sender: TObject );
var staction ,tmp: string;
begin
with TTobViewer(sender) do
  begin
  staction:='ACTION=CONSULTATION';
  if RTDroitModifTiers(AsString[ColIndex('AFF_TIERS'), CurrentRow]) then staction:='ACTION=MODIFICATION';
  if ((ColName[CurrentCol] = 'RPE_AUXILIAIRE') or (ColName[CurrentCol] = 'RPE_TIERS')
    or (ColName[CurrentCol] = 'AFF_TIERS') or (ColName[CurrentCol] = 'T_TIERS')
    or (ColName[CurrentCol] = 'T_LIBELLE') or (ColName[CurrentCol] = 'RaisonSociale')) then
       V_PGI.DispatchTT (28,taConsult ,AsString[ColIndex('RPE_AUXILIAIRE'), CurrentRow], '','')
  else if (ColName[CurrentCol] = 'RPE_NUMEROACTION') and  (AsInteger[ColIndex('RPE_NUMEROACTION'), CurrentRow] <> 0 )
  then V_PGI.DispatchTT (22,taConsult ,AsString[ColIndex('RPE_AUXILIAIRE'), CurrentRow]+';'+IntToStr(AsInteger[ColIndex('RPE_NUMEROACTION'), CurrentRow]), '','')
  else if (ColName[CurrentCol] = 'RPE_OPERATION') and (trim(AsString[ColIndex('RPE_OPERATION'), CurrentRow])<>'')
  then  V_PGI.DispatchTT (23,taConsult ,AsString[ColIndex('RPE_OPERATION'), CurrentRow], '','')

  else if ((ColName[CurrentCol] = 'RPE_PROJET') and (trim(AsString[ColIndex('RPE_PROJET'), CurrentRow])<>'')) or
       ((ColName[CurrentCol] = 'RPJ_PROJET') and (trim(AsString[ColIndex('RPJ_PROJET'), CurrentRow])<>'')) or
       (ColName[CurrentCol] = 'RPJ_LIBELLE')
  then V_PGI.DispatchTT (30,taConsult ,AsString[ColIndex('RPE_PROJET'), CurrentRow], '','')

  else if ((ColName[CurrentCol] = 'RPE_PERSPECTIVE') and (AsInteger[ColIndex('RPE_PERSPECTIVE'), CurrentRow]<>0)) or
          (ColName[CurrentCol] = 'AFF_PERSPECTIVE') and (AsInteger[ColIndex('AFF_PERSPECTIVE'), CurrentRow]<>0) or
          (ColName[CurrentCol] = 'RPE_LIBELLE')
  then V_PGI.DispatchTT (40,taConsult ,IntToStr(AsInteger[ColIndex('AFF_PERSPECTIVE'), CurrentRow]), '','')

  else if (ColName[CurrentCol] = 'T_SOCIETEGROUPE')  then  V_PGI.DispatchTT (8,taConsult ,AsString[ColIndex('T_SOCIETEGROUPE'), CurrentRow], '','')
  else if (ColName[CurrentCol] = 'T_PRESCRIPTEUR')  then  V_PGI.DispatchTT (8,taConsult ,AsString[ColIndex('T_PRESCRIPTEUR'), CurrentRow], '','')
  else
    begin
      if ctxscot in V_PGI.PGIContexte then tmp :='MISSION'  //mcd 06/10/2005 pas même fiche GI et GA
       else tmp := 'AFFAIRE';
      if AGLJaiLeDroitFiche(['AFFAIRE','ACTIOn=CONSULTATION','P'],3)
       then AglLanceFiche('AFF',tmp,'',AsString[ColIndex('AFF_AFFAIRE'), CurrentRow],'ACTION=CONSULTATION;MONOFICHE');
    end;
  end;
end;

procedure TOF_RTAFFAIRE_TV.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
// mcd 19/09/01 Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
// mcd 19/09/01 ne pas alimenter pour OK sélection aff Aff_:=THEdit(GetControl('AFF_AFFAIRE_'));
Aff0_:=THEdit(GetControl('AFF_AFFAIRE0_'));
Aff1_:=THEdit(GetControl('AFF_AFFAIRE1_'));
Aff2_:=THEdit(GetControl('AFF_AFFAIRE2_'));
Aff3_:=THEdit(GetControl('AFF_AFFAIRE3_'));
Aff4_:=THEdit(GetControl('AFF_AVENANT_'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
Tiers_:=THEdit(GetControl('AFF_TIERS_'));
end;
procedure TOF_RTAFFAIRE_TV.bSelectAff1Click(Sender: TObject);
begin
  EditAff0.Text:='P';
  SelectionAffaire (nil, EditAff, EditAff0, EditAff1, EditAff2, EditAff3, EditAff4, false, false, '', false, true, true);
end;
procedure TOF_RTAFFAIRE_TV.Select2Click(Sender: TObject);
begin
  EditAff0_.Text:='P';
  SelectionAffaire (nil,EditAff_,EditAff0_,EditAff1_,EditAff2_,EditAff3_,EditAff4_, false, false, '', false, true, true);
end;

Initialization
registerclasses([TOF_RTAFFAIRE_TV]);

end.
 