{***********UNITE*************************************************
Auteur  ...... : AB
Créé le ...... : 26/09/2002
Modifié le ... :   /  /
Description .. : Comparaison du Planning et du Réalisé
Mots clefs ... : TOF:AFPLANNINGCOMPAR  FICHE sta AFPLANNINGCOMPAR
*****************************************************************}
Unit UtofAFPlanningCompar;

Interface

Uses StdCtrls,Controls,Classes,Windows,sysutils,Graphics,
     UTOF,uTob,stat,HCtrls,HEnt1,UTobView,ParamSoc,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db,dbtables,FE_Main,
{$ENDIF}
     DicoAf,UTofAfBaseCodeAffaire,Afplanning,AFPlanningCst,ActiviteUtil,AFActivite,EntGC,UtilRessource;

Type
  TOF_AFPLANNINGCOMPAR = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    private
      fTobAffiche,fTobPlanning,fTobActivite,fTobRess : Tob;
      fTobViewer :TTobViewer;
      fDateDebPla,fDateFinPla,fDateDebReal,fDateFinReal:TDateTime;
      fUnite :String;
      procedure LoadPlanning;
      procedure LoadActivite;
      procedure LoadRessource;
      procedure MajRessource;
      procedure PutColQte (tobDest,tobSource:TOB;ChampQte:string);
      procedure PutResteAPlanifie (tobL:Tob);
      procedure PutResteAFaire (tobL:Tob);
      procedure ConvertQte(tobL:Tob ; ChampQte : string );
      procedure TVOnDblClickCell(Sender: TObject );
  end;

  Procedure AFLanceFiche_PlanningCompar;

const
      TexteMessage: array[1..3] of string = (
      {1}  'Vous n''avez pas de ressources pour cette sélection ',
      {2}  'Vous n''avez pas de ressources planifiées pour cette sélection ',
      {3}  'Vous n''avez pas d''activité pour cette sélection ');

Implementation

Procedure AFLanceFiche_PlanningCompar;
begin
  AGLLanceFiche ('AFF','AFPLANNINGCOMPAR ','','','') ;
end;

procedure TOF_AFPLANNINGCOMPAR.OnArgument (S : String ) ;
Begin
  Inherited;
  SetControlVisible('PAVANCE', false);
  SetControlVisible('FPresentations', true);
  SetControlVisible('PSQL', false);
  fTobAffiche := TOB.create('Les_resultats', nil, -1);
  fTobRess := TOB.create('Les_ressources', nil, -1);
  fTobPlanning := TOB.create('Les_Plannings', nil, -1);
  fTobActivite := TOB.create('Les_Activités', nil, -1);

  TFStat(Ecran).LaTOB :=  fTobAffiche;
  fTobViewer := TTobViewer(getcontrol('TV'));
  fTobViewer.OnDblClick:= TVOnDblClickCell ;
{$IFDEF EAGLCLIENT}
  setControlVisible('BIMPRIMER', false);
{$ENDIF}
  SetControlText ('UNITETEMPS',VH_GC.AFMesureActivite);
end;

procedure TOF_AFPLANNINGCOMPAR.OnLoad;
begin
  inherited;
  fTobAffiche.cleardetail;
  fTobRess.cleardetail;
  fTobPlanning.cleardetail;
  fTobActivite.cleardetail;
  fDateDebPla := StrToDate(GetControlText('DATEPLANIF'));
  fDateFinPla := StrToDate(GetControlText('DATEPLANIF_'));
  fDateDebReal := StrToDate(GetControlText('DATEREAL'));
  fDateFinReal := StrToDate(GetControlText('DATEREAL_'));
  fUnite := getcontroltext ('UNITETEMPS');  
  if (fDateDebReal = iDate1900) then
  begin
    fDateDebReal := fDateDebPla;
    SetControlText( 'DATEREAL' , GetControlText('DATEPLANIF'));
  end;
  if (fDateFinReal = iDate2099) then
  begin
    fDateFinReal := fDateFinPla;
    SetControlText( 'DATEREAL_' , GetControlText('DATEPLANIF_'));
  end;
end;

procedure TOF_AFPLANNINGCOMPAR.OnUpdate ;
begin
  Inherited ;
  if (Transactions(LoadRessource,3) = oeOk) then MajRessource;
end;

procedure TOF_AFPLANNINGCOMPAR.OnClose;
begin
  Inherited;
  fTobAffiche.cleardetail; fTobAffiche.free; fTobAffiche := nil;
  fTobRess.cleardetail; fTobRess.free; fTobRess := nil;
  fTobPlanning.cleardetail; fTobPlanning.Free; fTobPlanning := nil;
  fTobActivite.cleardetail; fTobActivite.Free; fTobActivite := nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : AB
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Remplir la tob virtuelle avec les données du planning
Mots clefs ... :
*****************************************************************}

procedure TOF_AFPLANNINGCOMPAR.LoadPlanning;
var
  vSt,vStCritere : String;
  vQR       : TQuery;
  i :integer;
begin
  vStCritere := trim(TFStat(Ecran).stsql);
  if (vStCritere = '') then vStCritere := ' WHERE '
  else vStCritere := FindEtReplace (vStCritere,'ATR_','APL_',true) + ' AND ';
  vSt := 'SELECT APL_AFFAIRE,APL_RESSOURCE,SUM(APL_QTEPLANIFUREF) AS PLANIFIE,SUM(APL_QTEREALUREF) AS REALISEPLA';
  vSt := vSt + ' FROM AFPLANNING ';
  vSt := vSt + ' LEFT JOIN TACHE ON APL_AFFAIRE=ATA_AFFAIRE AND APL_NUMEROTACHE=ATA_NUMEROTACHE ';
  vSt := vSt + ' LEFT JOIN RESSOURCE ON APL_RESSOURCE=ARS_RESSOURCE ';
  vSt := vSt + vStCritere + ' APL_RESSOURCE<>"" ';
  vSt := vSt + ' AND APL_DATEDEBPLA >= "'+USDateTime(fDateDebPla)+'"';
  vSt := vSt + ' AND APL_DATEFINPLA <= "'+USDateTime(fDateFinPla)+'"';
  vSt := vSt + ' GROUP BY APL_AFFAIRE,APL_RESSOURCE ';
  vSt := vSt + ' ORDER BY APL_AFFAIRE,APL_RESSOURCE ';
  vQR := OpenSql (vSt,True);

  Try
    if Not vQR.Eof then fTobPlanning.LoadDetailDB('LesPlannings','','',vQr,False,True)
  finally
    if vQr <> nil then Ferme(vQr);
  end;

  if (trim(fUnite) <> '') and(fUnite <> VH_GC.AFMESUREACTIVITE) then
  for i := 0 to fTobPlanning.detail.Count - 1 do
  begin
    ConvertQte( fTobPlanning.detail[i],'PLANIFIE');
    ConvertQte( fTobPlanning.detail[i],'REALISEPLA');
  end;
end;

{***********A.G.L.*******************************************************
Auteur  ...... : AB
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Remplir la tob virtuelle avec les données de l'activité
Mots clefs ... :
**************************************************************************}

procedure TOF_AFPLANNINGCOMPAR.LoadActivite;
var
  vSt,vStCritere : String;
  vQR       : TQuery;
  i : integer;
begin
  vStCritere := trim(TFStat(Ecran).stsql);
  if (vStCritere = '') then vStCritere := ' WHERE '
  else vStCritere := FindEtReplace (vStCritere,'ATR_','ACT_',true) + ' AND ';
  vSt := 'SELECT ACT_AFFAIRE,ACT_RESSOURCE,SUM(ACT_QTEUNITEREF) AS REALISEACT';
  vSt := vSt + ' FROM ACTIVITE ';
  vSt := vSt + ' LEFT JOIN TACHE ON ACT_AFFAIRE=ATA_AFFAIRE AND ACT_NUMEROTACHE=ATA_NUMEROTACHE ';
  vSt := vSt + ' LEFT JOIN RESSOURCE ON ACT_RESSOURCE=ARS_RESSOURCE ';
  vSt := vSt + vStCritere + ' ACT_RESSOURCE<>"" AND ACT_AFFAIRE<>"" ';
  if fDateDebReal <> iDate1900 then
  vSt := vSt + ' AND ACT_DATEACTIVITE >= "'+UsDateTime(fDateDebReal)+'"';
  if fDateFinReal <> iDate2099 then
  vSt := vSt + ' AND ACT_DATEACTIVITE <= "'+UsDateTime(fDateFinReal)+'"';
  vSt := vSt + ' AND ACT_ACTIVITEREPRIS <> "A" AND ACT_TYPEACTIVITE="REA" AND ACT_TYPEARTICLE="PRE"';
  vSt := vSt + ' AND ACT_ETATVISA="VIS"';
  vSt := vSt + ' GROUP BY ACT_AFFAIRE,ACT_RESSOURCE ';
  vSt := vSt + ' ORDER BY ACT_AFFAIRE,ACT_RESSOURCE ';
  vQR := OpenSql (vSt,True);

  Try
    if Not vQR.Eof then fTobActivite.LoadDetailDB('LesActivites','','',vQr,False,True)
  finally
    if vQr <> nil then Ferme(vQr);
  end;

  if (trim(fUnite) <> '') and(fUnite <> VH_GC.AFMESUREACTIVITE) then
  for i := 0 to fTobActivite.detail.Count - 1 do
    ConvertQte( fTobActivite.detail[i],'REALISEACT');
end;

{***********A.G.L.********************************************************************
Auteur  ...... : AB
Créé le ...... : 26/09/2002
Modifié le ... :   /  /
Description .. : TacheRessource groupées par ressource,affaire pour lignes d'affichage
Mots clefs ... :
***************************************************************************************}
procedure TOF_AFPLANNINGCOMPAR.LoadRessource;
var
  vSt,vStCritere : String;
  vQR : TQuery;
  i : integer;
begin
  LoadPlanning;
  LoadActivite;
  vStCritere := trim(TFStat(Ecran).stsql);
  if (vStCritere = '') then vStCritere := ' WHERE '
  else vStCritere := vStCritere + ' AND ';
  vSt := 'SELECT ATR_AFFAIRE,AFF_LIBELLE as DESCRIPTIF,ATR_RESSOURCE,ARS_LIBELLE AS RESSOURCE,';
  vSt := vSt + 'SUM(ATR_QTEINITUREF) AS QTEINITIALE,';
  vSt := vSt + '0.00 AS PLANIFIE,0.00 AS RESTEAPLANIFIE,0.00 AS REALISEPLA,0.00 AS REALISEACT ';
  if GetParamSoc('SO_AFGESTIONRAF') then
  vSt := vSt + ' ,SUM(ATR_QTEAPLANIFUREF) AS ECART,0.00 AS RESTEAFAIRE';
  vSt := vSt + ' FROM TACHERESSOURCE';
  vSt := vSt + ' LEFT JOIN AFFAIRE ON ATR_AFFAIRE=AFF_AFFAIRE ';
  vSt := vSt + ' LEFT JOIN RESSOURCE ON ATR_RESSOURCE=ARS_RESSOURCE ';
  vSt := vSt + ' LEFT JOIN TACHE ON ATR_AFFAIRE=ATA_AFFAIRE AND ATR_NUMEROTACHE=ATA_NUMEROTACHE ';
  vSt := vSt + vStCritere + ' ATR_RESSOURCE<>"" ';
  vSt := vSt + ' GROUP BY ATR_AFFAIRE,ATR_RESSOURCE,AFF_LIBELLE,ARS_LIBELLE';
  vSt := vSt + ' ORDER BY ATR_AFFAIRE,ATR_RESSOURCE,AFF_LIBELLE,ARS_LIBELLE ';
  vQR := OpenSql (vSt,True);
  Try
    if Not vQR.Eof then fTobRess.LoadDetailDB('AfficheRessource','','',vQr,False,True);
  finally
    if vQr <> nil then Ferme(vQr);
  end;

  if (trim(fUnite) <> '') and(fUnite <> VH_GC.AFMESUREACTIVITE) then
  for i := 0 to fTobRess.detail.Count - 1 do
  begin
    ConvertQte( fTobRess.detail[i],'QTEINITIALE');
    if GetParamSoc('SO_AFGESTIONRAF') then
      ConvertQte( fTobRess.detail[i],'ECART');
  end;
end;

procedure TOF_AFPLANNINGCOMPAR.MajRessource;
var vStRess,vStAff : String;
    i : Integer;
    vTobAff,vTC : Tob;
begin
  if (fTobRess.detail.Count=0) then
    PGIInfoAf( textemessage[1], Ecran.Caption)
  else if (fTobPlanning.detail.count=0) then
    PGIInfoAf( textemessage[2], Ecran.Caption)
  else if (fTobActivite.detail.count=0) then
    PGIInfoAf( textemessage[3], Ecran.Caption);

  for i := 0 to fTobRess.detail.Count - 1 do
  begin
    vTobAff := fTobRess.Detail[i];
    vStAff  := vTobAff.getvalue('ATR_AFFAIRE');
    vStRess := vTobAff.getvalue('ATR_RESSOURCE');
    vTC := fTobPlanning.FindFirst(['APL_AFFAIRE','APL_RESSOURCE'],[vStAff,vStRess],TRUE) ;
    if vTC<>Nil then
    begin
      PutColQte (vTobAff,vTC,'PLANIFIE');
      PutColQte (vTobAff,vTC,'REALISEPLA');
    end;
    vTC := fTobActivite.FindFirst(['ACT_AFFAIRE','ACT_RESSOURCE'],[vStAff,vStRess],TRUE) ;
    if vTC<>Nil then PutColQte (vTobAff,vTC,'REALISEACT');
    PutResteAPlanifie (vTobAff);
    if GetParamSoc('So_AFGESTIONRAF') then
      PutResteAFaire (vTobAff);
  end;
  for i := fTobRess.Detail.Count-1 downto 0 do
  begin
    if (fTobRess.detail[i]<>nil) and
    ((fTobRess.detail[i].getvalue('PLANIFIE')>0) or (fTobRess.detail[i].getvalue('REALISEACT')>0)) THEN
      fTobRess.detail[i].changeparent(fTobAffiche, -1);
  end;
end;

procedure TOF_AFPLANNINGCOMPAR.PutColQte (tobDest,tobSource:TOB;ChampQte:string);
begin
  tobDest.putvalue(ChampQte,tobDest.getvalue(ChampQte) + tobSource.getvalue(ChampQte) );
end;

{***********************************************************************************
Calcul de la quantité du reste à Planifier
Reste : Prévu - Planifié
RESTE : QTEINITIALE-PLANIFIE
***********************************************************************************}
procedure TOF_AFPLANNINGCOMPAR.PutResteAPlanifie (tobL:Tob);
var vQte :double;
begin
  vQte := tobL.getvalue('QTEINITIALE')-tobL.getvalue('PLANIFIE');
  if (vQte > 0.0) then tobL.putvalue ('RESTEAPLANIFIE',vQte);
end;

{***********************************************************************************
Calcul de la quantité du reste à faire RAF
Reste : Prévu + Ecart - Réalisé Activité
RESTE : QTEINITIALE+ECART-REALISEACT
***********************************************************************************}
procedure TOF_AFPLANNINGCOMPAR.PutResteAFaire (tobL:Tob);
var vQte :double;
begin
  vQte := tobL.getvalue('QTEINITIALE') + tobL.getvalue('ECART') - tobL.getvalue('REALISEACT');
  if (vQte > 0.0) then tobL.putvalue ('RESTEAFAIRE',vQte);
end;

{***********************************************************************************
Convertie la quantité de l'unité de temps de référence vers l'unité saisie
***********************************************************************************}
procedure TOF_AFPLANNINGCOMPAR.ConvertQte(tobL:Tob ; ChampQte : string );
begin
  if (trim(fUnite) <> '') and(fUnite <> VH_GC.AFMESUREACTIVITE) then
  tobL.putvalue(ChampQte, ConversionUnite(VH_GC.AFMESUREACTIVITE,fUnite ,tobL.getvalue(ChampQte)));
end;

procedure TOF_AFPLANNINGCOMPAR.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
Begin
  Aff   := THEdit(GetControl('ATR_AFFAIRE'));
  Aff1  := THEdit(GetControl('ATR_AFFAIRE1'));
  Aff2  := THEdit(GetControl('ATR_AFFAIRE2'));
  Aff3  := THEdit(GetControl('ATR_AFFAIRE3'));
  Aff4  := THEdit(GetControl('ATR_AVENANT'));
  Tiers := THEdit(GetControl('ATR_TIERS'));
End;

procedure TOF_AFPLANNINGCOMPAR.TVOnDblClickCell(Sender: TObject );
var vStArg :string;
begin
with fTobViewer do
  begin
    if (ColName[CurrentCol]='ATR_AFFAIRE') or (ColName[CurrentCol]='DESCRIPTIF')
    or (ColName[CurrentCol]='QTEINITIALE') or (ColName[CurrentCol]='QTEAPLANIFIER') then
    begin
//      AFLanceFiche_Affaire (AsString[ColIndex('ATR_AFFAIRE'),CurrentRow],'STATUT:AFF;ETAT:ENC');
      V_PGI.DispatchTT( 5,taModif,AsString[ColIndex('ATR_AFFAIRE'),CurrentRow], '', 'STATUT:AFF;ETAT:ENC') ;
    end
    else if (ColName[CurrentCol]='REALISEACT') then
    begin
      vStArg := 'RESSOURCE='+AsString[ColIndex('ATR_RESSOURCE'),CurrentRow];
      vStArg := vStArg + ';AFFAIRE='+AsString[ColIndex('ATR_AFFAIRE'),CurrentRow];
      vStArg := vStArg + ';DATEDEB='+DateToStr(fDateDebReal);
      vStArg := vStArg + ';DATEFIN='+DateToStr(fDateFinReal);
      vStArg := vStArg + ';TYPEARTICLE=PRE';
      vStArg := vStArg + ';FACT=N|F';
      AGLLanceFiche ('AFF','AFACTIVITECON_MUL','','',vStArg);
  //  AFCreerActiviteModale (tsaRess,tacTemps,'REA',AsString[ColIndex('ATR_RESSOURCE'),CurrentRow],AsString[ColIndex('ATR_AFFAIRE'),CurrentRow],'',fDateDebReal)
    end
    else
    begin
//      vStArg := ' APL_AFFAIRE = "'+AsString[ColIndex('ATR_AFFAIRE'),CurrentRow]+'"';
      ExecPlanning('153202',DateToStr(fDateDebPla),vStArg, '', '-');
    end                                                   
  end;
end;

Initialization
  registerclasses ([TOF_AFPLANNINGCOMPAR]) ;
end.



