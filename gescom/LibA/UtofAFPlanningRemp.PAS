{***********UNITE*******************************************************
Auteur  ...... : AB
Créé le ...... : 27/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFPLANNINGREMP
Remplacement d'une ressource dans le planning et les taches ressources
Mots clefs ... : TOF;AFPLANNINGREMP
***********************************************************************}
Unit UtofAFPlanningRemp;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
   MaineAGL,
{$ELSE}
   dbTables, db,FE_Main,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Lookup,
     UTOB,UTOF,UtofAfBaseCodeAffaire,AffaireUtil,DicoAf,
     UtilTaches,
     UtilRessource, uafo_ressource;

Type
 TOF_AFPLANNINGREMP = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnLoad; override ;
    procedure OnUpdate ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
  private
    fStAffaire,fStClient,fStFamilleTache,fStRessource2 :string;
    fDtDebut,fDtFin :TdateTime;
    fTobTacheRess1,fTobTacheRess2 :TOB;
    fBDefinitif : Boolean;
    fED_RESSOURCE1,fED_NUMEROTACHE : THEdit;
    procedure ChangeAffaire(Sender: TObject);
    procedure ChangeTache (Sender: TObject);
    procedure RechRessource1(Sender: TObject);
    procedure ChangeRessource (Sender: TObject);
    procedure RemplaceRessource;
    procedure RemplaceRessPlanning;
    procedure CreationTacheRessource;
    procedure ValoriseRessource;
    procedure ValoriseRessTache;
    procedure ValoriseRessPlanning;
  end;
const
  TexteMsg: array[1..8] of string 	= (
    {1}        'Vous devez saisir le code de la ressource à remplacer.',
    {2}        'Vous devez saisir le code de la ressource de remplacement.',
    {3}        'Le code de la ressource n''existe pas.',
    {4}        'Confirmez-vous le remplacement de la ressource ',
    {5}        'Le remplacement de la ressource s''est effectué correctement.',
    {6}        'Le remplacement de la ressource ne s''est pas effectué correctement.',
    {7}        'La valorisation du planning et des tâches s''est effectuée correctement.',
    {8}        'La valorisation du planning et des tâches ne s''est pas effectuée correctement.');
                                         

procedure AFLanceFiche_PlanningRemp;

Implementation

procedure AFLanceFiche_PlanningRemp;
begin
  AGLLanceFiche('AFF','AFPLANNINGREMP','', '','');
end;

procedure TOF_AFPLANNINGREMP.OnLoad ;
begin
  Inherited ;
  fDtDebut := iDate1900;   fDtFin := iDate2099;
  THEdit(getcontrol('AFF_AFFAIRE')).OnChange := ChangeAffaire;
  fED_RESSOURCE1 := THEdit(getcontrol('RESSOURCE1'));
  fED_RESSOURCE1.OnElipsisClick := RechRessource1;
  fED_RESSOURCE1.OnExit := ChangeRessource;
  THEdit(getcontrol('RESSOURCE2')).OnExit := ChangeRessource;
  fED_NUMEROTACHE := THEdit(getcontrol('NUMEROTACHE'));
  fED_NUMEROTACHE.OnChange := ChangeTache;
end ;

procedure TOF_AFPLANNINGREMP.OnUpdate ;
begin
  Inherited ;
  fBDefinitif := TRadioButton(GetControl('RB_DEFINITIF')).checked;
  if IsValidDate(getcontroltext ('DATEPLANIF')) then
    fDtDebut := strtodate(getcontroltext ('DATEPLANIF'));
  if IsValidDate(getcontroltext ('DATEPLANIF_')) then
    fDtFin   := strtodate(getcontroltext ('DATEPLANIF_'));
  fstAffaire := getcontroltext ('AFF_AFFAIRE');
  fStClient := getcontroltext ('AFF_TIERS');
  fStFamilleTache := THValCombobox(getcontrol('FAMILLETACHE')).value;
  if (trim(fED_RESSOURCE1.text) = '') then
  begin
    PGIBoxAF (TexteMsg[1],ecran.caption);
    fED_RESSOURCE1.SetFocus;
    exit;
  end;
  fStRessource2 := getcontroltext ('RESSOURCE2');
  if (trim(fStRessource2) = '') then
  begin
    PGIBoxAF (TexteMsg[2],ecran.caption);
    SetFocusControl('RESSOURCE2');
    exit;
  end;
  if (PGIAskAF(TexteMsg[4]+fED_RESSOURCE1.text+' par la ressource '+fStRessource2+' ?',ecran.caption)= mrYes) then
  begin
    fTobTacheRess1:= Tob.create('Les TACHERESSOURCE1', nil, -1);
    fTobTacheRess2:= Tob.create('Les TACHERESSOURCE2', nil, -1);

    if (Transactions(RemplaceRessource,3) = oeOk) then PGIInfoAF (TexteMsg[5],ecran.caption)
    else PGIBoxAF (TexteMsg[6],ecran.caption);

    if (TCheckBox(GetControl('BVALORISE')).Checked) then  //Valorisation
    begin
      if (Transactions(ValoriseRessource,3) = oeOk) then PGIInfoAF (TexteMsg[7],ecran.caption)
      else PGIBoxAF (TexteMsg[8],ecran.caption);
    end;

    fTobTacheRess1.cleardetail;  fTobTacheRess1.free;
    fTobTacheRess2.cleardetail;  fTobTacheRess2.free;
  end;
end ;

procedure TOF_AFPLANNINGREMP.RemplaceRessource;
begin
  RemplaceRessPlanning;
  CreationTacheRessource;
end;

procedure TOF_AFPLANNINGREMP.RemplaceRessPlanning;
var vSt :string;
begin
  vSt := ' UPDATE AFPLANNING SET APL_RESSOURCE = "'+fStRessource2+'"';
  vSt := vSt + ' WHERE APL_RESSOURCE = "'+fED_RESSOURCE1.text+'"';
  if (trim (fstAffaire) <> '') then
  vSt := vSt + ' AND APL_AFFAIRE = "'+fstAffaire+'"';
  if (trim (fStClient) <> '') then
  vSt := vSt + ' AND APL_TIERS = "'+fStClient+'"';
  if (fDtDebut <> iDate1900) then
  vSt := vSt + ' AND APL_DATEDEBPLA >= "'+UsDateTime(fDtDebut)+'"';
  if (fDtFin <> iDate2099) then
  vSt := vSt + ' AND APL_DATEFINPLA <= "'+UsDateTime(fDtFin)+'"';

  if (trim (fED_NUMEROTACHE.text) = '') and (trim (fStFamilleTache) <> '') then
  begin
    vSt := vSt + ' AND APL_NUMEROTACHE IN (SELECT ATA_NUMEROTACHE ';
    vSt := vSt + ' FROM TACHE WHERE ATA_FAMILLETACHE = "'+fStFamilleTache+'"';
    if (trim (fstAffaire) <> '') then
    vSt := vSt + ' AND ATA_AFFAIRE = "'+fstAffaire+'"';
    if (trim (fStClient) <> '') then
    vSt := vSt + ' AND ATA_TIERS = "'+fStClient+'"';
    vSt := vSt + ')';
  end else if (trim (fED_NUMEROTACHE.text) <> '') then
  begin
    vSt := vSt + ' AND APL_NUMEROTACHE = '+fED_NUMEROTACHE.text;
  end;
  ExecuteSql(vSt);
end;

procedure TOF_AFPLANNINGREMP.CreationTacheRessource;
var vSt :string;
    vQR : TQuery;
    vTob :TOB;
    i :integer;
begin

  vSt := 'SELECT * FROM TACHERESSOURCE'; // Besoin de tous les champs pour duplication enregistrement
  if (trim (fED_NUMEROTACHE.text) = '') and (trim (fStFamilleTache) <> '') then
  vSt := vSt + ' LEFT JOIN TACHE ON ATR_AFFAIRE = ATA_AFFAIRE AND ATR_NUMEROTACHE = ATA_NUMEROTACHE';
  vSt := vSt + ' WHERE ATR_RESSOURCE =  "' + fED_RESSOURCE1.text + '"';
  if (trim (fstAffaire) <> '') then
  vSt := vSt + ' AND ATR_AFFAIRE =  "' + fStAffaire + '"';
  if (trim (fStClient) <> '') then
  vSt := vSt + ' AND ATR_TIERS = "'+fStClient+'"';
  if (trim (fED_NUMEROTACHE.text) <> '') then
  vSt := vSt + ' AND ATR_NUMEROTACHE =  ' + fED_NUMEROTACHE.text
  else if (trim (fStFamilleTache) <> '') then
  vSt := vSt + ' AND ATA_FAMILLETACHE = "'+fStFamilleTache+'"';

  vQr := nil;
  Try
    vQR := OpenSql(vSt,True);
    if Not vQR.Eof then
    begin
      fTobTacheRess1.LoadDetailDB('TACHERESSOURCE','','', vQR, False, True);
      for i:=0 to fTobTacheRess1.detail.count-1 do
      begin
        vTob := Tob.create ('TACHERESSOURCE',fTobTacheRess2,-1);
        vTob.dupliquer (fTobTacheRess1.detail[i],FALSE,TRUE,TRUE) ;
        vTob.putvalue('ATR_RESSOURCE',fStRessource2);
        if fBDefinitif then
        begin
          vTob.putvalue('ATR_STATUTRES','ACT');
          fTobTacheRess1.detail[i].putvalue('ATR_STATUTRES','INA');
        end else vTob.putvalue('ATR_STATUTRES','TEM');
      end;
      if (fTobTacheRess2.InsertOrUpdateDB(true)) and fBDefinitif then
        fTobTacheRess1.UpdateDB(True);
    end;
  finally
    if vQr <> nil then Ferme(vQr);
  end;
end;

procedure TOF_AFPLANNINGREMP.ValoriseRessource;
begin
  ValoriseRessTache;
  ValoriseRessPlanning;
end;

procedure TOF_AFPLANNINGREMP.ValoriseRessTache;
var
  vTobTaches      : TOB;
  vSt             : string;
  vQR             : TQuery;
  i               :integer;
  vTOBAffaires    : Tob;
  vTobArticles    : Tob;
  vAFOAssistants  : TAFO_Ressources;

begin

  vTobTaches := Tob.create('TACHES', nil, -1);
  vSt := 'SELECT * FROM TACHE WHERE ATA_NUMEROTACHE IN ';
  vSt := vSt + ' ( SELECT ATR_NUMEROTACHE  FROM TACHERESSOURCE';
  vSt := vSt + ' WHERE ATR_RESSOURCE =  "' + fStRessource2 + '"';

  if (trim (fstAffaire) <> '') then
    vSt := vSt + ' AND ATR_AFFAIRE =  "' + fStAffaire + '"';
  if (trim (fStClient) <> '') then
    vSt := vSt + ' AND ATR_TIERS = "'+fStClient+'"';
  if (trim (fED_NUMEROTACHE.text) <> '') then
    vSt := vSt + ' AND ATR_NUMEROTACHE =  '+ fED_NUMEROTACHE.text;
  vSt := vSt + ' )';
  if (trim (fstAffaire) <> '') then
    vSt := vSt + ' AND ATA_AFFAIRE =  "' + fStAffaire + '"';
  if (trim (fStClient) <> '') then
    vSt := vSt + ' AND ATA_TIERS = "'+fStClient+'"';
  if (trim (fStFamilleTache) <> '') then
    vSt := vSt + ' AND ATA_FAMILLETACHE = "'+fStFamilleTache+'"';

  vQr := nil;
  try
    vQR := OpenSql(vSt,True);
    If (not vQR.EOF) then
    begin
      vTobTaches.LoadDetailDB('TACHE','','', vQR,True);
      OrganiseTacheRessource (vTobTaches, fTobTacheRess2);
                                             
      vAFOAssistants  := TAFO_Ressources.Create;
      vTOBAffaires    := TOB.Create('Les Affaires',Nil,-1) ;
      vTOBArticles    := TOB.Create('Les Articles',Nil,-1) ;
      try
        for i:=0 to vTobTaches.detail.count-1 do
          Valorisation(vTobTaches.detail[i], 'ATA', vAFOAssistants, vTOBAffaires, vTobArticles);
      finally
        vTOBAffaires.Free;
        vTobArticles.Free;
        vAFOAssistants.Free;
      end;
    end;
  finally
    Ferme(vQR);
    vTobTaches.cleardetail;
    vTobTaches.free;
  End;
end;

procedure TOF_AFPLANNINGREMP.ValoriseRessPlanning;
var
  vTobPlanning    : TOB;
  vSt             : string;
  vQR             : TQuery;
  i               : integer;
  vTOBAffaires    : Tob;
  vTobArticles    : Tob;
  vAFOAssistants  : TAFO_Ressources;

begin
  vTobPlanning := Tob.create('Les plannings', nil, -1);

  vSt := ' SELECT * FROM AFPLANNING ';
  vSt := vSt + ' WHERE APL_RESSOURCE = "'+fStRessource2+'"';
  if (trim (fstAffaire) <> '') then
    vSt := vSt + ' AND APL_AFFAIRE = "'+fstAffaire+'"';
  if (trim (fStClient) <> '') then
    vSt := vSt + ' AND APL_TIERS = "'+fStClient+'"';    
  if (fDtDebut <> iDate1900) then
    vSt := vSt + ' AND APL_DATEDEBPLA >= "'+UsDateTime(fDtDebut)+'"';
  if (fDtFin <> iDate2099) then
    vSt := vSt + ' AND APL_DATEFINPLA <= "'+UsDateTime(fDtFin)+'"';
  if (trim (fED_NUMEROTACHE.text) <> '') then
    vSt := vSt + ' AND APL_NUMEROTACHE = '+fED_NUMEROTACHE.text
  else if (trim (fStFamilleTache) <> '') then
  begin
    vSt := vSt + ' AND APL_NUMEROTACHE IN (SELECT ATA_NUMEROTACHE ';
    vSt := vSt + ' FROM TACHE WHERE ATA_FAMILLETACHE = "'+fStFamilleTache+'"';
    if (trim (fstAffaire) <> '') then
      vSt := vSt + ' AND ATA_AFFAIRE = "'+fstAffaire+'"';
    if (trim (fStClient) <> '') then
      vSt := vSt + ' AND ATA_TIERS = "'+fStClient+'"';
    vSt := vSt + ')';
  end;

  vQr := nil;
  try
    vQR := OpenSql(vSt,True);
    If (not vQR.EOF) then
    begin
                                                  
      vAFOAssistants  := TAFO_Ressources.Create;
      vTOBAffaires    := TOB.Create('Les Affaires',Nil,-1);
      vTOBArticles    := TOB.Create('les Articles',Nil,-1) ;
      try
        vTobPlanning.LoadDetailDB('AFPLANNING','','', vQR,True);
        For i:=0 to vTobPlanning.detail.count-1 do
          Valorisation(vTobPlanning.detail[i], 'APL', vAFOAssistants, vTOBAffaires, vTobArticles);
        vTobPlanning.UpdateDB(true);
      finally
        vTOBAffaires.Free;
        vTobArticles.Free;
        vAFOAssistants.Free;
      end;
    end;
  finally
    Ferme(vQR);                     
    vTobPlanning.cleardetail;
    vTobPlanning.free;
  end;
end;

procedure TOF_AFPLANNINGREMP.RechRessource1(Sender: TObject);
var vSt :string;
BEGIN
  fstAffaire := getcontroltext ('AFF_AFFAIRE');
  fstClient := getcontroltext ('AFF_TIERS');
  if (trim(fstAffaire) <> '') or (trim(fstClient) <> '') then
  begin
    vSt := ' ARS_RESSOURCE IN (SELECT ATR_RESSOURCE FROM TACHERESSOURCE';
    vSt := vSt + ' WHERE ATR_STATUTRES = "ACT"';
    if (trim(fstAffaire) <> '') then
      vSt := vSt + ' AND ATR_AFFAIRE="'+fstAffaire+'"'
    else if (trim(fstClient) <> '') then
      vSt := vSt + ' AND ATR_TIERS="'+fstClient+'"';
    if (trim(fED_NUMEROTACHE.text) <> '') then
      vSt := vSt + ' AND ATR_NUMEROTACHE='+ fED_NUMEROTACHE.text ;
    vSt := vSt +')';
    LookupList(fED_RESSOURCE1,'Ressources','RESSOURCE','ARS_RESSOURCE','ARS_LIBELLE',vSt,'',True,24,'',tlDefault );
  end else GetRessourceMul (fED_RESSOURCE1, '', '', '');
end;

procedure TOF_AFPLANNINGREMP.ChangeAffaire(Sender: TObject);
var vSt :string;
    vQR : Tquery;
begin
  fstAffaire := getcontroltext ('AFF_AFFAIRE');
  if (trim(fstAffaire) <> '') then
  begin
    fED_NUMEROTACHE.plus := ' AND ATA_AFFAIRE = "' + fStAffaire + '"';;
    vSt := 'SELECT AFF_LIBELLE, T_LIBELLE FROM TIERS, AFFAIRE ';
    vSt := vSt + 'WHERE AFF_AFFAIRE = "' + fStAffaire + '"';
    vSt := vSt + 'AND AFF_TIERS = T_TIERS';
    vQr := nil;
    Try
      vQR := OpenSql(vSt, True);
      if Not vQR.Eof then
        begin
          SetControlText('LIBAFFAIRE', vQR.FindField('AFF_LIBELLE').AsString);
          SetControlText('LIBCLIENT', vQR.FindField('T_LIBELLE').AsString);
        end;
    finally
      if vQR <> nil then ferme(vQR);
    end;
  end else
  begin
    SetControlText('LIBAFFAIRE', '');
    SetControlText('LIBCLIENT', '');
    fED_NUMEROTACHE.plus := '';
  end;
end;

procedure TOF_AFPLANNINGREMP.ChangeTache(Sender: TObject);
var vSt :string;
    vQR : Tquery;
begin
  if (trim(fED_NUMEROTACHE.text) <> '') then
  begin
    vSt := 'SELECT ATA_FAMILLETACHE FROM TACHE ';
    vSt := vSt + ' WHERE ATA_NUMEROTACHE = "' + fED_NUMEROTACHE.text + '"';
    vQr := nil;
    Try
      vQR := OpenSql(vSt, True);
      if Not vQR.Eof then SetControlText('FAMILLETACHE', vQR.FindField('ATA_FAMILLETACHE').AsString);
    finally
      if vQR <> nil then ferme(vQR);
    end;
  end;
end;

procedure TOF_AFPLANNINGREMP.ChangeRessource(Sender: TObject);
var QQ : TQuery ;
    vCodeRess,vRess : string;
begin
  vCodeRess := trim(THEDIT(Sender).text);
  vRess := THEDIT(Sender).name;
  SetControlText('FON'+vRess,'');    // Fonction actuelle de la ressource
  SetControlText('LIB'+vRess, '');   // Nom de la ressource
  if  vCodeRess = '' then exit;
  QQ:=OpenSQL('SELECT ARS_LIBELLE,ARS_FONCTION1 FROM RESSOURCE WHERE ARS_RESSOURCE="'+ vCodeRess+'"',false);
  if not QQ.EOF then
  begin
    SetControlText('LIB'+vRess, QQ.findField('ARS_LIBELLE').asString) ;
    SetControlText('FON'+vRess,RechDom('AFLFONCTION',QQ.findField('ARS_FONCTION1').asString,False));
  end
  else begin
    SetFocusControl(vRess);
    PGIBoxAF (TexteMsg[3],ecran.caption);
  end;
  ferme(QQ);
end;

procedure TOF_AFPLANNINGREMP.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
  Aff:=THEdit(GetControl('AFF_AFFAIRE'));
  Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
  Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
  Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
  Aff4:=THEdit(GetControl('AFF_AVENANT'));
  Tiers:=THEdit(GetControl('AFF_TIERS'));
end;

Initialization
  registerclasses ( [ TOF_AFPLANNINGREMP ] );
end.
