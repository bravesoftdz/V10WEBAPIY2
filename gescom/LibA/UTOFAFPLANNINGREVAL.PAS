{***********UNITE*************************************************
Auteur  ...... : AB
Créé le ...... : 06/02/2003
Modifié le ... :   /  /
Description .. : Revalorisation du planning par affaire et par tâche
Fiche Mul .... : AFPLANNINGREVAL - AFTACHEREVAL_MUL
Mots clefs ... : TOF;AFPLANNINGREVAL
*****************************************************************}
Unit UtofAfPlanningReval;
     
Interface

Uses StdCtrls, Controls, Classes,  forms,Graphics, sysutils,  ComCtrls,
{$IFDEF EAGLCLIENT}
   Maineagl,eMul,
{$ELSE}
   dbTables,FE_Main,Mul,
{$EndIF}
   HCtrls,HEnt1,ed_tools,Utob,Utof,UtilMulTrt,DicoAF,utofAfBaseCodeAffaire,
   uafo_ressource,UtilTaches,TraducAffaire,AfUtilArticle ;

Type
  TOF_AFPLANNINGREVAL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnUpdate ; override ;
    procedure OnLoad ; override ;
    procedure OnArgument(stArgument : String ) ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    private
      fDateDebPla,fDateFinPla : TDateTime;
      fTobPlanning : TOB;
      fSelection :string;
      procedure BOuvrirOnClick (Sender: TObject);
      procedure ChargePlanning (pAffaire,pNumeroTache :string);
      function  RevalorisationPlanning :boolean;
  end ;

Procedure AFLanceFicheAFPlanningRevalAffaire;
Procedure AFLanceFicheAFPlanningRevalTache ;

const

// libellés des messages
TexteMessage: array[1..5] of string 	= (
    {1}  'Pas de planning à revaloriser.'
    {2} ,'%d lignes de planning à revaloriser, voulez-vous continuer ?'
    {3} ,'La revalorisation du planning a été effectuée avec succès.'
    {4} ,'La revalorisation du planning ne s''est pas effectuée.'
    {5} ,'Revalorisation du planning en cours...'
    );

Implementation

Procedure AFLanceFicheAFPlanningRevalAffaire ;
begin
  AGLLanceFiche('AFF','AFPLANNINGREVAL','','','AFF');
end;

Procedure AFLanceFicheAFPlanningRevalTache ;
begin
  AGLLanceFiche('AFF','AFTACHEREVAL_MUL','','','ATA');
end;

procedure TOF_AFPLANNINGREVAL.OnArgument(stArgument : String );
begin
  Inherited ;
  TFMul(Ecran).BOuvrir.OnClick:= BOuvrirOnClick;
  fSelection := stArgument;
  if fSelection = 'ATA' then
    THMultiValComboBox(GetControl('ATA_TYPEARTICLE')).plus:=PlusTypeArticle;

  {$IFDEF CCS3}
  if (Ecran.Name = 'AFPLANNINGREVAL') then
  begin
    if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
    if (getcontrol('PSTAT') <> Nil) then SetControlVisible ('PSTAT', False);
  end;
  if (Ecran.Name = 'AFTACHEREVAL_MUL') then
  begin
    if (getcontrol('PTEXTELIBRE') <> Nil) then SetControlVisible ('PTEXTELIBRE', False);
    if (getcontrol('PZONELIBRES') <> Nil) then SetControlVisible ('PZONELIBRES', False);
  end;
  {$ENDIF}
end;

procedure TOF_AFPLANNINGREVAL.OnUpdate ;
begin
  Inherited ;
{$IFDEF EAGLCLIENT}
  TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
  TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
end ;

procedure TOF_AFPLANNINGREVAL.BOuvrirOnClick(Sender: TObject);
var vTobEnregs :TOB;
    vRet : T_ErrTraiteEnreg;
    i : integer;
begin
  fDateDebPla := StrToDate(GetControlText('DATEREVAL'));
  fDateFinPla := StrToDate(GetControlText('DATEREVAL_'));
  vTobEnregs := TOB.Create('Les Enregs',Nil,-1);
  fTobPlanning := TOB.Create('Le planning',Nil,-1);
  try
    SourisSablier;
    if fselection = 'AFF' then
    begin
      vRet := TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFFAIRE', vTobEnregs, True);
      if vRet <> tteOK then exit ;
      for i:=0 to vTobEnregs.detail.count-1 do
        ChargePlanning (vTobEnregs.detail[i].getvalue('AFF_AFFAIRE'),'');
    end
    else if fselection = 'ATA' then
    begin
      vRet := TraiteEnregMulListe (TFMul(Ecran), 'ATA_NUMEROTACHE','TACHE', vTobEnregs, True);
      if vRet <> tteOK then exit ;
      for i:=0 to vTobEnregs.detail.count-1 do
        ChargePlanning (vTobEnregs.detail[i].getvalue('ATA_AFFAIRE'),vTobEnregs.detail[i].getvalue('ATA_NUMEROTACHE'));
    end;

    if fTobPlanning.detail.count = 0 then
    begin
      PgiInfoAf(TexteMessage[1], Ecran.Caption);  Exit;
    end;
    if (PGIAskAF (format(TexteMessage[2],[fTobPlanning.detail.count]),Ecran.Caption) <> mrYes) then Exit;
    if RevalorisationPlanning then
      PgiInfoAf(TexteMessage[3], Ecran.Caption)
    else PgiInfoAf(TexteMessage[4], Ecran.Caption);

  finally
    vTobEnregs.free;
    fTobPlanning.free;
    SourisNormale;
  end;
end;

procedure TOF_AFPLANNINGREVAL.ChargePlanning (pAffaire,pNumeroTache :string);
var vStReq :string;
    QQ : TQuery;
begin
  vStReq := ' SELECT * FROM AFPLANNING ';
  vStReq := vStReq + ' WHERE APL_AFFAIRE = "' + pAffaire+'"';
  if (fselection = 'ATA') and (trim(pNumeroTache) <> '') then
    vStReq := vStReq + ' AND APL_NUMEROTACHE = '+ pNumeroTache;
  if (fDateDebPla > iDate1900) then
  vStReq := vStReq + ' AND APL_DATEDEBPLA >= "'+UsDateTime(fDateDebPla)+'"';
  if (fDateFinPla < iDate2099) then
  vStReq := vStReq + ' AND APL_DATEFINPLA <= "'+UsDateTime(fDateFinPla)+'"';

  QQ := OpenSQL (vStReq, True);
  if Not QQ.EOF then fTobPlanning.LoadDetailDB('AFPLANNING','','', QQ,True,True);
  Ferme(QQ);
end;

function TOF_AFPLANNINGREVAL.RevalorisationPlanning : boolean;
var vTOBArticles,vTOBAffaires : TOB;
    vAFOAssistants : TAFO_Ressources;
    i : integer;
begin
  Result := false;
  vAFOAssistants  := TAFO_Ressources.Create;
  vTOBArticles := TOB.Create('Les Articles',Nil,-1);
  vTOBAffaires := TOB.Create('Les Affaires',Nil,-1);
  InitMoveProgressForm (nil,Ecran.Caption,'',fTobPlanning.detail.count,TRUE,TRUE) ;

  try
    fTobPlanning.SetAllModifie (false);
    for i := 0 to fTobPlanning.detail.count-1 do
    begin
      if Not MoveCurProgressForm(TexteMessage[5]) then Exit ;
      Valorisation (fTobPlanning.detail[i], 'APL', vAFOAssistants, vTOBAffaires, vTOBArticles);
    end;
    Result := fTobPlanning.UpdateDB(true);

  finally
    vAFOAssistants.free;
    vTOBArticles.free;
    vTOBAffaires.free;
    FiniMoveProgressForm ;
  end;
end;

procedure TOF_AFPLANNINGREVAL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin
  inherited;
  if (TFMul(Ecran).name='AFPLANNINGREVAL') then
  begin
    Aff   := THEdit(GetControl('AFF_AFFAIRE'));
    Aff1  := THEdit(GetControl('AFF_AFFAIRE1'));
    Aff2 := THEdit(GetControl('AFF_AFFAIRE2'));
    Aff3  := THEdit(GetControl('AFF_AFFAIRE3'));
    Aff4 := THEdit(GetControl('AFF_AVENANT'));
    Tiers := THEdit(GetControl('AFF_TIERS'));
    // affaire de référence pour recherche
    Aff_  := THEdit(GetControl('AFF_AFFAIREREF'));
    Aff1_ := THEdit(GetControl('AFFAIREREF1'));
    Aff2_ := THEdit(GetControl('AFFAIREREF2'));
    Aff3_ := THEdit(GetControl('AFFAIREREF3'));
    Aff4_ := THEdit(GetControl('AFFAIREREF4'));
  end
  else if (TFMul(Ecran).name='AFTACHEREVAL_MUL') then
  begin
    Aff   := THEdit(GetControl('ATA_AFFAIRE'));
    Aff1  := THEdit(GetControl('ATA_AFFAIRE1'));
    Aff2  := THEdit(GetControl('ATA_AFFAIRE2'));
    Aff3  := THEdit(GetControl('ATA_AFFAIRE3'));
    Aff4  := THEdit(GetControl('ATA_AVENANT'));
    Tiers := THEdit(GetControl('ATA_TIERS'));
  end;
end;

procedure TOF_AFPLANNINGREVAL.OnLoad;
begin
  inherited;
end;

Initialization
  registerclasses ( [ TOF_AFPLANNINGREVAL ] ) ;
end.


