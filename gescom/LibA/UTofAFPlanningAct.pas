unit UTofAFPlanningAct;

interface

uses  StdCtrls,Controls,Classes,M3FP,sysutils,ComCtrls
      ,HTB97,hqry,utob,HCtrls,HEnt1,HMsgBox,HStatus,ed_tools
{$IFDEF EAGLCLIENT}
      ,MaineAGL ,Emul
{$ELSE}
      ,db,dbtables,FE_Main,HDB ,Mul
{$ENDIF}
      ,Ent1,EntGC,DicoAF,AffaireUtil,utofAfBaseCodeAffaire,AffaireRegroupeUtil
      ,TraducAffaire,FactActivite,ActiviteUtil,Paramsoc
      ;

Type
   TOF_AFPlanningAct = Class (TOF_AFBASECODEAFFAIRE)
      procedure OnArgument(stArgument : String ); override;
      procedure OnUpdate; override;
      procedure OnLoad; override;
      procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
      procedure BOuvrirOnClick(Sender : TObject);
      function  GenerePlanningActivite:boolean;
      procedure GenereActivite;
      function  PlanningToAct: boolean;
      function  PlanningMaj (FromTOB: TOB): boolean;
   private
     fStWhere : String;
     fTobPlaMaj,fTobActi,fTobCleAct,fTobPlan :Tob;
   End;

Procedure AFLanceFiche_GenPlanningAct;

const

fMaxAffaire = 20000;
// libellés des messages
TexteMessage: array[1..8] of string 	= (
    {1}  'Aucune affaire sélectionnée.'
    {2} ,' affaires sélectionnées. La sélection est trop importante, vous devez la restreindre.'
    {3} ,' affaires sélectionnées. Confirmez-vous la génération du planning dans l''activité ? '
    {4} ,'Pas de réalisé du planning à générer dans l''activité jusqu''au  '
    {5} ,'Le réalisé du planning a été généré dans l''activité jusqu''au '
    {6} ,'La date fin de tâches saisie est postérieure à la date fin d''activité des paramètres.#10#13 L''activité ne sera donc générée que jusqu''au '
    {7} ,'La génération du planning dans l''activité a été interrompue.'
    {8} ,'La génération du planning dans l''activité a échouée pour l''affaire : '
        );

implementation

Procedure AFLanceFiche_GenPlanningAct;
begin
  AGLLanceFiche ('AFF','AFPLANNINGACT','','','');
end;

procedure TOF_AFPlanningAct.OnArgument(stArgument : String);
var
  vStTmp    : String;
Begin
  Inherited;
  // traitement des arguments
  vStTmp:= (Trim(ReadTokenSt(stArgument)));
  // et les sous affaires
  if not(GereSousAffaire) then
  begin
    SetcontrolVisible('TAFF_AFFAIREREF',False);
    SetcontrolVisible('AFFAIREREF1',False); SetcontrolVisible('AFFAIREREF2',False);
    SetcontrolVisible('AFFAIREREF3',False); SetcontrolVisible('AFFAIREREF4',False);
    SetcontrolVisible('AFF_ISAFFAIREREF',False);
    SetcontrolVisible('BSELECTAFF2',False);
  end;

{$IFDEF EAGLCLIENT}
  TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
  TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
  TFMul(Ecran).BOuvrir.OnClick := BOuvrirOnClick ;
  {$IFDEF CCS3}
  if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
  if (getcontrol('PSTAT') <> Nil) then SetControlVisible ('PSTAT', False);
  {$ENDIF}
End;

procedure TOF_AFPlanningAct.OnUpdate;
Begin
  // Gestion repositionnement auto sur l'affaire en cours si sortie rapide / bug par prg ( ou Eagl)
  inherited;
  if Not (VH_GC.CleAffaire.Co2Visible) then SetControlText('AFF_AFFAIRE2','');
  if Not (VH_GC.CleAffaire.Co3Visible) then SetControlText('AFF_AFFAIRE3','');
End;

procedure TOF_AFPlanningAct.OnLoad;
Var Affaire,Affaire0,Affaire1,Affaire2,Affaire3,Avenant : string;
begin
  inherited;
  if GereSousAffaire then
  begin
    if (GetControlText('AFFAIREREF1')='') then SetControlText('AFF_AFFAIREREF','')
    else 
    begin
      Affaire0 := 'A';
      Affaire1 := GetControlText('AFFAIREREF1'); Affaire2 := GetControlText('AFFAIREREF2');
      Affaire3 := GetControlText('AFFAIREREF3'); Avenant := GetControlText('AFFAIREREF4');
      Affaire:=CodeAffaireRegroupe(Affaire0, Affaire1  ,Affaire2 ,Affaire3,Avenant, taModif, false,false,false);
      if not ExisteAffaire(Affaire,'') then SetControlText('AFF_AFFAIREREF',Trim(Copy(Affaire,1,15)));
    end;
  end;
end;

procedure TOF_AFPlanningAct.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('AFF_AFFAIRE'));
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1')); Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3')); Aff4:=THEdit(GetControl('AFF_AVENANT'));
  Tiers:=THEdit(GetControl('AFF_TIERS'));

  // affaire de référence pour recherche
  Aff_:=THEdit(GetControl('AFF_AFFAIREREF'));
  Aff1_:=THEdit(GetControl('AFFAIREREF1'));
  Aff2_:=THEdit(GetControl('AFFAIREREF2'));Aff3_:=THEdit(GetControl('AFFAIREREF3'));
  Aff4_:=THEdit(GetControl('AFFAIREREF4'));
end;

procedure TOF_AFPlanningAct.BOuvrirOnClick(Sender : TObject);
var
  i,iNbAffaire : Integer;
  QQ : TQuery;
{$IFDEF EAGLCLIENT}
  fListe: THGrid;
{$ELSE}
  fListe : THDBGrid;
{$ENDIF}
begin
  fListe := TFMul(Ecran).FListe;
  fStWhere := Copy(RecupWhereCritere(TPageControl(GetControl('PAGES'))),6,length(RecupWhereCritere(TPageControl(GetControl('PAGES')))) );
  // on ajoute aux criteres la liste des affaires sélectionnées
  try
    if (not fListe.AllSelected) and (fListe.nbSelected <> 0) then
      begin
        for i := 0 to fListe.nbSelected -1 do
          begin
            fListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
            TFmul(Ecran).Q.TQ.Seek(fListe.Row-1) ;
{$ENDIF}
            if (i = 0) and (fStWhere <> '') then fStWhere := fStWhere + ' AND ( '
            else if (i <> 0) and (fStWhere <> '') then fStWhere := fStWhere + ' OR ';
            fStWhere := fStWhere + 'AFF_AFFAIRE = "' + TFmul(Ecran).Q.Findfield('AFF_AFFAIRE').AsSTring + '"';
          end;
        if fStWhere <> '' then fStWhere := fStWhere + ')';
      end;
   except
   on E:Exception do
    begin
     MessageAlerte('Erreur lors du chargement du planning' +#10#13#10#13 + E.message ) ;
    end; // on
   end; // try
  iNbAffaire:=0;
  if fListe.AllSelected then
  begin
    QQ := nil;
    try
      QQ := OpenSql ('SELECT COUNT(AFF_AFFAIRE) AS NBAFF FROM AFFAIRE WHERE ' + fStWhere, True);
      if not QQ.Eof then iNbAffaire := QQ.FindField('NBAFF').AsInteger;
    finally
      Ferme(QQ);
    end;
  end
  else iNbAffaire := fListe.NbSelected;

  if (iNbAffaire = 0) then
    PGIInfoAf( textemessage[1], Ecran.Caption)
  else if (iNbAffaire > fMaxAffaire) then
    PGIInfoAf( inttostr(iNbAffaire) + textemessage[2], Ecran.Caption)
  else if (PGIAskAF (inttostr(iNbAffaire) + textemessage[3], Ecran.Caption) = mrYes) then
  begin
    EnableControls(Ecran, False );
    Try
      GenerePlanningActivite;
    Finally
      EnableControls(Ecran, True);
      if FListe.AllSelected then FListe.AllSelected:=False
      else FListe.ClearSelected;
      TFMul(Ecran).bSelectAll.Down := False ;
    end;
  end;
end;

//******************************************************************************
//**** Déversement du planning des affaires sélectionnées dans l'activité ******
//******************************************************************************

Function TOF_AFPlanningAct.GenerePlanningActivite:boolean;
var S       : string;
    Q       : Tquery;
    vDtFin,vDateDebutAct,vDateFinAct  : TDateTime;
    vStDateFin : string;
begin
  result := false;
  vStDateFin := GetControlText ('FINPLANNING');
  vDtFin := StrToDate(vStDateFin);

  // control de la date de fin de generation de l'activite
  IntervalleDatesActivite(vDateDebutAct, vDateFinAct);
  if vDateFinAct < vDtFin then
  begin
    vStDateFin := datetostr(vDateFinAct);
    PGIInfoAf( textemessage[6] + vStDateFin, Ecran.Caption);
    vDtFin := vDateFinAct;
    SetControlText ('FINPLANNING',vStDateFin);
  end;

  fTobPlan := TOB.Create('Les_Plannings',Nil,-1) ;
  fTobActi := TOB.Create('Les_Activités',Nil,-1) ;
  fTobCleAct := TOB.Create ('liste cles act' , Nil,-1);
  fTobPlaMaj :=  TOB.Create('Maj_Plannings',Nil,-1) ;

  S :=    ' SELECT AFPLANNING.*,ARS_TYPERESSOURCE,ARS_FONCTION1 ';
  S := S+ ' ,ATR_COMPETENCE1,ATR_COMPETENCE2,ATR_COMPETENCE3, ATA_LIBELLETACHE1 ';
  S := S+ ' FROM AFPLANNING,AFFAIRE,RESSOURCE,TACHERESSOURCE, TACHE ';
  S := S+ ' WHERE APL_DATEDEBREAL >= "'+ UsDateTime(vDateDebutAct)+'"';
  S := S+ ' AND APL_DATEFINREAL <= "'+ UsDateTime(vDtFin)+'"';
  S := S+ ' AND APL_AFFAIRE = AFF_AFFAIRE';
  S := S+ ' AND APL_RESSOURCE = ARS_RESSOURCE';
  S := S+ ' AND APL_AFFAIRE = ATR_AFFAIRE';
  S := S+ ' AND APL_RESSOURCE = ATR_RESSOURCE';
  S := S+ ' AND APL_NUMEROTACHE = ATR_NUMEROTACHE';
  S := S+ ' AND APL_ACTIVITEGENERE = "-"';
  S := S+ ' AND APL_QTEREALISE > 0';
  S := S+ ' AND ATA_AFFAIRE = APL_AFFAIRE ';
  S := S+ ' AND ATA_NUMEROTACHE = APL_NUMEROTACHE ';
  S := S+ ' AND APL_TERMINE = "X" ';

  if fStWhere <> '' then  S := S+ ' AND '+ fStWhere;
  S := S+ ' ORDER BY APL_AFFAIRE,APL_RESSOURCE,APL_DATEDEBPLA ';
  MoveCur (False);
  Q := nil;
  Try
    Q:=OpenSql(S,True);
    if Not Q.Eof then fTobPlan.LoadDetailDB('LignesPLANNING','','',Q,False,True);
    Ferme(Q);

    InitMoveProgressForm (Ecran,'Génération de l''activité...','', fTobPlan.detail.Count, true, true) ;

    if fTobPlan.detail.count > 0 then
    begin
      if PlanningToAct then
        PGIInfoAf(textemessage[5] + vStDateFin, Ecran.Caption)
      else
        PGIInfoAf(textemessage[7], Ecran.Caption)
    end
    else PGIInfoAf(textemessage[4] + vStDateFin, Ecran.Caption);

  Finally
    fTobPlan.cleardetail; fTobPlan.Free;
    fTobActi.cleardetail; fTobActi.free;
    fTobCleAct.cleardetail; fTobCleAct.free;
    fTobPlaMaj.cleardetail; fTobPlaMaj.free;
    FiniMove;
    FiniMoveProgressForm;
  end;
END;

procedure TOF_AFPlanningAct.GenereActivite;
var i,NumLigneUnique : integer;
    vCodeAffaire : string;
begin
  vCodeAffaire := fTobActi.detail[0].GetValue('ACT_AFFAIRE');
  // PL le 15/04/03 : modif clé activité : le planning est déversé dans les négatifs
  NumLigneUnique := ProchainMoinsNumLigneUniqueActivite ('REA', vCodeAffaire);
  for i := 0 to fTobActi.detail.Count-1 do
    fTobActi.detail[i].PutValue('ACT_NUMLIGNEUNIQUE', NumLigneUnique - i);
  if fTobActi.Detail.count > 0 then fTobActi.InsertDB(Nil);
  if fTobPlaMaj.Detail.count > 0 then fTobPlaMaj.InsertOrUpdateDB(False);
end;

function TOF_AFPlanningAct.PlanningToAct:boolean;
Var NumLigne,i : integer;
    TobDetAct,TobLigne : TOB;
    vCodeAffaire : string;
    io :TIOErr;
BEGIN
  Result:=true;
  NumLigne := 10000;
  if (fTobPlan = Nil) or (fTobActi = Nil) or (fTobPlan.detail.Count=0) then Exit;
  vCodeAffaire := fTobPlan.detail[0].GetValue('APL_AFFAIRE');
  for i := 0 to fTobPlan.detail.Count-1 do
  begin
    TobLigne := fTobPlan.detail[i];
    if (vCodeAffaire <> TobLigne.GetValue('APL_AFFAIRE')) then
    begin
      io := Transactions(GenereActivite,3); // Génération du planning dans l'activité par affaire
      if (io <> oeOk) then
      begin
        PGIInfoAf(textemessage[8] + vCodeAffaire, Ecran.Caption);
        result:= false;
        break;
      end;
      vCodeAffaire := TobLigne.GetValue('APL_AFFAIRE');
      fTobActi.cleardetail;
      fTobPlaMaj.cleardetail;
      NumLigne := 10000;
    end;

    TobDetAct := TOB.Create ('ACTIVITE', fTobActi,-1);
    TobDetAct.InitValeurs;
    TobDetAct.PutValue('ACT_TYPEACTIVITE', 'REA');
    TobDetAct.PutValue('ACT_AFFAIRE', TobLigne.GetValue('APL_AFFAIRE'));
    TobDetAct.PutValue('ACT_AFFAIRE0', TobLigne.GetValue('APL_AFFAIRE0'));
    TobDetAct.PutValue('ACT_AFFAIRE1', TobLigne.GetValue('APL_AFFAIRE1'));
    TobDetAct.PutValue('ACT_AFFAIRE2', TobLigne.GetValue('APL_AFFAIRE2'));
    TobDetAct.PutValue('ACT_AFFAIRE3', TobLigne.GetValue('APL_AFFAIRE3'));
    TobDetAct.PutValue('ACT_AVENANT', TobLigne.GetValue('APL_AVENANT'));
    TobDetAct.PutValue('ACT_TIERS', TobLigne.GetValue('APL_TIERS'));
    TobDetAct.PutValue('ACT_NUMAPPREC', 0);
    TobDetAct.PutValue('ACT_NUMEROTACHE', TobLigne.GetValue('APL_NUMEROTACHE'));
    TobDetAct.PutValue('ACT_ACTORIGINE', 'APL');
    TobDetAct.PutValue('ACT_RESSOURCE', TobLigne.GetValue('APL_RESSOURCE'));
    if (TobLigne.GetValue('APL_RESSOURCE') <> '') then
    begin
      TobDetAct.PutValue('ACT_TYPERESSOURCE', TobLigne.GetValue('ARS_TYPERESSOURCE'));
      TobDetAct.PutValue('ACT_FONCTIONRES', TobLigne.GetValue('APL_FONCTION'));
    end;
    TobDetAct.PutValue('ACT_DATEACTIVITE', TobLigne.GetValue('APL_DATEDEBREAL'));
    TobDetAct.PutValue('ACT_FOLIO', 01);
    TobDetAct.PutValue('ACT_PERIODE', GetPeriode(TobLigne.GetValue('APL_DATEDEBREAL')));
    TobDetAct.PutValue('ACT_SEMAINE', NumSemaine(TobLigne.GetValue('APL_DATEDEBREAL')));

    TobDetAct.PutValue('ACT_TYPEARTICLE', TobLigne.GetValue('APL_TYPEARTICLE'));
    TobDetAct.PutValue('ACT_ARTICLE', TobLigne.GetValue('APL_ARTICLE'));
    TobDetAct.PutValue('ACT_CODEARTICLE', TobLigne.GetValue('APL_CODEARTICLE'));

    // libellé
    TobDetAct.PutValue('ACT_LIBELLE', TobLigne.GetValue('ATA_LIBELLETACHE1'));
    TobDetAct.PutValue('ACT_UNITE', TobLigne.GetValue('APL_UNITETEMPS')); // ou stock ???
    TobDetAct.PutValue('ACT_QTE', TobLigne.GetValue('APL_QTEREALISE'));
    TobDetAct.PutValue('ACT_QTEUNITEREF', TobLigne.GetValue('APL_QTEREALUREF'));
    TobDetAct.PutValue('ACT_HEUREDEBUT', TobLigne.GetValue('APL_HEUREDEBREAL'));
    TobDetAct.PutValue('ACT_HEUREFIN', TobLigne.GetValue('APL_HEUREFINREAL'));
    TobDetAct.PutValue('ACT_UNITEFAC', TobLigne.GetValue('APL_UNITETEMPS'));

    TobDetAct.PutValue('ACT_QTEFAC', TobLigne.GetValue('APL_QTEREALISE'));

    TobDetAct.PutValue('ACT_PUPR', TobLigne.GetValue('APL_PUPR'));
    TobDetAct.PutValue('ACT_PUVENTE', TobLigne.GetValue('APL_PUVENTEHT'));
    TobDetAct.PutValue('ACT_PUVENTEDEV', TobLigne.GetValue('APL_PUVENTEDEVHT'));
    TobDetAct.PutValue('ACT_DEVISE', TobLigne.GetValue('APL_DEVISE'));
    TobDetAct.PutValue('ACT_TOTPR', TobLigne.GetValue('APL_INITPTPR'));
    TobDetAct.PutValue('ACT_TOTPRCHARGE', TobLigne.GetValue('APL_INITPTPR'));
    TobDetAct.PutValue('ACT_TOTPRCHINDI', TobLigne.GetValue('APL_INITPTPR'));
    TobDetAct.PutValue('ACT_TOTVENTE', TobLigne.GetValue('APL_INITPTVENTEHT'));
    TobDetAct.PutValue('ACT_TOTVENTEDEV', TobLigne.GetValue('APL_INITPTVTDEVHT'));

    TobDetAct.PutValue('ACT_ACTIVITEREPRIS', TobLigne.GetValue('APL_ACTIVITEREPRIS'));
    if (trim (TobDetAct.GetValue('ACT_ACTIVITEREPRIS'))='') then
      TobDetAct.PutValue('ACT_ACTIVITEREPRIS','F');
    TobDetAct.PutValue('ACT_COMPETENCE1', TobLigne.GetValue('ATR_COMPETENCE1'));
    TobDetAct.PutValue('ACT_COMPETENCE2', TobLigne.GetValue('ATR_COMPETENCE2'));
    TobDetAct.PutValue('ACT_COMPETENCE3', TobLigne.GetValue('ATR_COMPETENCE1'));
    TobDetAct.PutValue('ACT_CREATEUR', V_PGI.User);
    TobDetAct.PutValue('ACT_DESCRIPTIF', TobLigne.GetValue('APL_DESCRIPTIF'));
    TobDetAct.PutValue('ACT_ACTIVITEEFFECT','X');

    //flagé la ligne planning si l'activité a été généré
    PlanningMaj(TobLigne);

    // Gestion automatique du visa sur l'activité
    if (GetParamSoc('SO_AFVISAACTIVITE')=false) then
    begin
      TobDetAct.PutValue('ACT_ETATVISA', 'VIS');
      TobDetAct.PutValue('ACT_VISEUR', V_PGI.User);
      TobDetAct.PutValue('ACT_DATEVISA', NowH);
    end
    else
    begin
      TobDetAct.PutValue('ACT_ETATVISA', 'ATT');
      TobDetAct.PutValue('ACT_DATEVISA', idate1900);
    end;

    // Gestion automatique du visa de facturation sur l'activité
    if (GetParamSoc('SO_AFAPPPOINT')=false) then
    begin
      TobDetAct.PutValue('ACT_ETATVISAFAC', 'VIS');
      TobDetAct.PutValue('ACT_VISEURFAC', V_PGI.User);
      TobDetAct.PutValue('ACT_DATEVISAFAC', NowH);
    end
    else
    begin
     TobDetAct.PutValue('ACT_ETATVISAFAC', 'ATT');
     TobDetAct.PutValue('ACT_DATEVISAFAC', idate1900);
    end;

    TobDetAct.PutValue('ACT_NUMLIGNE', NumLigne);
    inc(NumLigne);

    if not MoveCurProgressForm ('Mise à jour Activité en cours...') then
    begin
      result:= false;
      break;
    end;
  end;

  if Result then
  begin
    io := Transactions(GenereActivite,3);
    if (io <> oeOk) then
    begin
      PGIInfoAf(textemessage[8] + vCodeAffaire, Ecran.Caption);
      Result := False;
    end;
  end;
END;

function TOF_AFPlanningAct.PlanningMaj (FromTOB: TOB): boolean;
Var ToTOB : Tob;
    Q : Tquery;
begin
  Result := false;
  if (FromTOB = Nil) or (fTobPlaMaj = Nil) then Exit;
  ToTOB := TOB.Create ('AFPLANNING', fTobPlaMaj,-1);
  ToTOB.InitValeurs;
  Q := OpenSql ('SELECT * FROM AFPLANNING WHERE APL_AFFAIRE="'+FromTOB.getvalue('APL_AFFAIRE')+'" AND APL_NUMEROLIGNE="'+IntToStr(FromTOB.getvalue('APL_NUMEROLIGNE'))+'"',False);
  ToTOB.selectDB('',Q,False);
  Ferme(Q);
  ToTOB.PutValue('APL_ACTIVITEGENERE','X');
end;

Initialization
  registerclasses([TOF_AFPlanningAct]);
end.
