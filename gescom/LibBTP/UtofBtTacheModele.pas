{***********UNITE*************************************************
Auteur  ...... : AB
Créé le ...... : 19/02/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTTACHEMODELE_MUL ()
Mots clefs ... : TOF;BTTACHEMODELE
*****************************************************************}
unit UTofBTTacheModele;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFDEF EAGLCLIENT}
  eMul,
  Maineagl,
{$ELSE}
  db,
{$IFNDEF DBXPRESS}dbtables {BDE},
{$ELSE}uDbxDataSet,
{$ENDIF}mul,
  FE_Main,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  menus,
  HMsgBox,
  HTB97,
  ed_tools,
  M3FP,
  UTOF,
  UTOB,
  LookUp,
  paramsoc,
  GereTobInterne,
  EntGC,
  AfUtilArticle,
  UtilMulTrt,
  UtilTaches,
  UtilGC,
  AffaireRegroupeUtil,
  UTofAfBaseCodeAffaire,
  DicoBTP,
  ConfidentAffaire,
  UAFO_Ressource;

type
  TOF_BTTACHEMODELE = class(TOF)
    procedure OnArgument(S: string); override;

    procedure OnLoad; override;
  private
    bMulti: Boolean;
    fTOBModeles: TOB;

    procedure BOuvrir1OnClick(Sender: TObject);

  end;

  TOF_BTTACHEMODELE_GEN = class(TOF_AFBASECODEAFFAIRE)
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;

  private
    bMultiModele: Boolean;
    fTOBTaches, fTOBAffaires, fTOBModeles, FtobJourModele: TOB;
    procedure SelectionModeleTache(Sender: TObject);
    procedure BOuvrirOnClick(Sender: TObject);
    procedure MultiModeleClick(Sender: TObject);
    function CreationTaches: boolean;
    procedure CreationUneTache(pTOBAffaire, pTOBModele: TOB; pNumTache: integer);
    procedure TOBCopieChamp(FromTOB, ToTOB: TOB);
    procedure AddRemoveItemFromPopup(stPopup, stItem: string; bVisible: boolean);
  end;

procedure AFLanceFicheAFTacheModele_Mul;
procedure AFLanceFicheAFTacheModele_Gen;
function AFGetTacheModele(pEDTacheModele: THCritMaskEdit; pLequel: string): string;

implementation
const
  // libellés des messages
  TexteMessage: array[1..12] of string = (
    {1}  'Vous n''avez pas sélectionné les modèles de tâche.'
    {2}, 'Vous avez sélectionné %d tâches à créer dans %d affaires, voulez-vous continuer ?'
    {3}, 'La création des tâches a été effectuée avec succès.'
    {4}, 'La création des tâches ne s''est pas effectuée.'
    {5}, 'Création des tâches en cours...'
    {6}, 'Enregistrement des tâches en cours...'
    {7}, 'Valorisation des tâches en cours...'
    {8}, ''
    {9}, ''
    {10},'Sélection des modèles de tâches'
    {11},'Recherche des modèles de tâches'
    {12},'Vous n''avez pas sélectionné de modèle.'
    );


procedure AFLanceFicheAFTacheModele_Mul;
begin
  AGLLanceFiche('BTP', 'BTTACHEMODELE_MUL', '', '', '');
end;

procedure AFLanceFicheAFTacheModele_Gen;
begin
  AGLLanceFiche('BTP', 'BTTACHEGEN_MUL', '', '', '');
end;

procedure TOF_BTTACHEMODELE.OnArgument(S: string);
begin
  inherited;

  SetControlVisible('bOuvrir1', False);

  if S = 'MULTI' then
  begin
    bMulti := true;
{$IFDEF EAGLCLIENT}
    SetControlProperty('Fliste','MultiSelect',true);
{$ELSE}
    SetControlProperty('Fliste','Multiselection',true);
{$ENDIF}
    SetControlVisible('bOuvrir1', true);
    SetControlVisible('bOuvrir', False);
    setcontrolVisible('BSELECTALL', true);
    setcontrolVisible('BINSERT', false);
    setcontrolText('MODE', 'MULTI');
    Ecran.Caption := TraduitGA(TexteMessage[10]);		//Sélection des modèles de tâches
    TToolBarButton97(GetControl('BOUVRIR1')).OnClick := BOuvrir1OnClick;
  end
  else
    if S = 'RECH' then
  begin
    setcontrolVisible('BSELECTALL', false);
    setcontrolVisible('BINSERT', false);
    setcontrolText('MODE', 'RECH');
    Ecran.Caption := TraduitGA(TexteMEssage[11]);		//Recherche des modèles de tâches
  end;
{$IFDEF CCS3}
  if (getcontrol('PTEXTELIBRE') <> nil) then
    SetControlVisible('PTEXTELIBRE', False);
  if (getcontrol('PZONE') <> nil) then
    SetControlVisible('PZONE', False);
{$ENDIF}
end;

procedure TOF_BTTACHEMODELE.OnLoad;
begin
  inherited;

  SetControlProperty('AFM_TYPEARTICLE', 'Plus', PlusTypeArticle);

  // traduction champs libres
  GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'AFM_LIBRETACHE', 10, '_');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'AFM_DATELIBRE', 3, '_');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'AFM_CHARLIBRE', 3, '_');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'AFM_VALLIBRE', 3, '_');
  GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'AFM_BOOLLIBRE', 3, '_');
end;

procedure TOF_BTTACHEMODELE.BOuvrir1OnClick(Sender: TObject);
Var StReq   : String;
    StWhere : String;
    StTob   : String;
begin

  fTOBModeles := MaTobInterne('Liste ModeleTaches');

  fTOBModeles.cleardetail;

  StReq   := 'SELECT * FROM AFMODELETACHE';
  StWhere := 'AFM_MODELETACHE';
  StTob   := 'AFMODELETACHE';

  TraiteEnregMulTable(TFMul(Ecran), StReq, StWhere, StTob, 'AFM_MODELETACHE', 'AFMODELETACHE', fTOBModeles, True);

  Ecran.Close;

end;

{****************************************************************
Auteur  ...... : AB
Créé le ...... : 19/02/2003
Modifié le ... :   /  /
Description .. : TOF de la FICHE : AFTACHEMODELE_MUL ()
Mots clefs ... : TOF;AFTACHEMODELE_GEN
Menu...........: Affectation de tâches à partir des modèles
*****************************************************************}

procedure TOF_BTTACHEMODELE_GEN.NomsChampsAffaire(var Aff, Aff0, Aff1,
  Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff := THEdit(GetControl('AFF_AFFAIRE'));
  Aff1 := THEdit(GetControl('AFF_AFFAIRE1'));
  Aff2 := THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3 := THEdit(GetControl('AFF_AFFAIRE3'));
  Aff4 := THEdit(GetControl('AFF_AVENANT'));
  Tiers := THEdit(GetControl('AFF_TIERS'));
end;

procedure TOF_BTTACHEMODELE_GEN.OnArgument(S: string);
begin
  inherited;
  //BVoirStatut := true;
  if not (ModifAffaireAutorise) then
    SetControlText('XXAction', 'ACTION=CONSULTATION');
  fTOBModeles := MaTobInterne('Liste ModeleTaches');
  TFMul(Ecran).BOuvrir.OnClick := BOuvrirOnClick;
  TToolBarButton97(GetControl('BSELECTMOD')).onClick := SelectionModeleTache;
  TCheckBox(GetControl('MULTIMODELE')).onClick := MultiModeleClick;
{$IFDEF CCS3}
  if (getcontrol('PSTAT') <> nil) then
    SetControlVisible('PSTAT', False);
  if (getcontrol('PZONE') <> nil) then
    SetControlVisible('PZONE', False);
{$ENDIF}
  if not (GereSousAffaire) then
    SetcontrolVisible('AFF_ISAFFAIREREF', False);
end;

procedure TOF_BTTACHEMODELE_GEN.OnClose;
begin
  inherited;
  if (VH_GC.AFTobInterne <> nil) then
    DetruitMaTobInterne('Liste ModeleTaches');
  FtobJourModele.free;
end;

procedure TOF_BTTACHEMODELE_GEN.AddRemoveItemFromPopup(stPopup, stItem: string; bVisible: boolean); //PCS
var
  pop: TPopupMenu;
  i: integer;
  st: string;
begin
  pop := TPopupMenu(GetControl(stPopup));
  if pop <> nil then
    for i := 0 to pop.items.count - 1 do
    begin
      st := uppercase(pop.items[i].name);
      if st = stItem then
      begin
        pop.items[i].visible := bVisible;
        break;
      end;
    end;
end;

procedure TOF_BTTACHEMODELE_GEN.SelectionModeleTache(Sender: TObject);
var
  CBModele: THValcomboBox;
  i: integer;
begin

  AGLLanceFiche('AFF', 'AFTACHEMODELE_MUL', '', '', 'MULTI');

  CBModele := THValcomboBox(GetControl('MODELETACHEMULTI'));
  CBModele.Items.clear;
  CBModele.Values.clear;
  CBModele.ItemIndex := -1;

  if (fTOBModeles <> nil) and (fTOBModeles.detail.Count > 0) then
     begin
     for i := 0 to fTOBModeles.detail.count - 1 do
         begin
         CBModele.Items.add(fTOBModeles.detail[i].GetValue('AFM_LIBELLETACHE1'));
         CBModele.Values.add(fTOBModeles.detail[i].GetValue('AFM_MODELETACHE'));
         end;
     CBModele.ItemIndex := 0;
     end;

end;

procedure TOF_BTTACHEMODELE_GEN.MultiModeleClick(Sender: TObject);
var
  CBModele: THValcomboBox;
begin
  bMultiModele := (GetControlText('MULTIMODELE') = 'X');
  SetControlVisible('BSELECTMOD', (GetControlText('MULTIMODELE') = 'X'));
  SetControlVisible('MODELETACHE', (GetControlText('MULTIMODELE') = '-'));
  SetControlVisible('MODELETACHEMULTI', (GetControlText('MULTIMODELE') = 'X'));

  if GetControlText('MULTIMODELE') <> 'X' then
     begin //mcd 28/01/04 effacement zone de sélection
     CBModele := THValcomboBox(GetControl('MODELETACHEMULTI'));
     CBModele.Items.clear;
     CBModele.Values.clear;
     CBModele.ItemIndex := -1;
     CBModele.text := '';
     AddRemoveItemFromPopup('POPMENU', 'MNMODELE', true);
     end
  else
     AddRemoveItemFromPopup('POPMENU', 'MNMODELE', False);

end;

procedure TOF_BTTACHEMODELE_GEN.BOuvrirOnClick(Sender: TObject);
var
  vSelect, Sql: string;
  TobDetModele, TobDet: Tob;
  ii: integer;
begin
  try
    SourisSablier;
    fTOBAffaires := Tob.create('Liste Affaires', nil, -1);
    vSelect := 'SELECT * FROM AFFAIRE';
    TraiteEnregMulTable(TFMul(Ecran), vSelect, 'AFF_AFFAIRE', 'AFFAIRE', 'AFF_AFFAIRE', 'AFFAIRE', fTOBAffaires, True);
    if fTOBAffaires.detail.count = 0 then
      Exit;
    if not bMultiModele then
    begin
      if GetControltext('MODELETACHE') = '' then
      begin
        PgiInfoAf(TexteMessage[12], Ecran.Caption);  //Vous n''avez pas sélectionné de modèle
        Exit;
      end;
      fTOBModeles.cleardetail; //mcd 28/01/04 sinon ajoute tout le temps des enrgt ...
      TobDetModele := TOB.Create('AFMODELETACHE', fTOBModeles, -1);
      TobDetModele.InitValeurs;
      TobDetModele.SelectDB('"' + GetControltext('MODELETACHE') + '"', nil);
      if FtobJourModele <> nil then
        FtobJourModele.free;
      FtobJourModele := Tob.create('AFTACHEJOUR', nil, -1);
      Sql := 'SELECT * FROM AfTacheJour WHERE ATJ_TYPEJOUR="MOD" AND ATJ_MODELETACHE="'
        + GetControltext('MODELETACHE') + '"';
      FTobJourModele.LoadDetailDBFromSQL('AFTACHEJOUR', SQL, True);
    end
    else // cas sélection multiples
      if fTOBModeles.detail.count = 0 then
    begin
      PgiInfoAf(TexteMessage[1], Ecran.Caption);
      Exit;
    end
    else
    begin // il faut charge la table tachejour
      if FtobJourModele <> nil then
        FtobJourModele.free;
      FtobJourModele := Tob.create('AFTACHEJOUR', nil, -1);
      Sql := 'SELECT * FROM AfTacheJour WHERE ATJ_TYPEJOUR="MOD" AND ATJ_MODELETACHE in (';
      for ii := 0 to FtobModeles.detail.count - 1 do
      begin
        Tobdet := FtobModeles.detail[ii];
        if ii <> 0 then
          Sql := Sql + '",';
        Sql := Sql + '"' + Tobdet.getvalue('AFM_MODELETACHE');
      end;
      Sql := Sql + '")';
      FTobJourModele.LoadDetailDBFromSQL('AFTACHEJOUR', SQL, True);
    end;

    if (PGIAskAF(format(TexteMessage[2], [fTOBModeles.detail.count, fTOBAffaires.detail.count]), Ecran.Caption) <> mrYes) then
      Exit;

    if CreationTaches then
      PgiInfoAf(TexteMessage[3], Ecran.Caption)
    else
      PgiInfoAf(TexteMessage[4], Ecran.Caption);

  finally
    SourisNormale;
    fTOBAffaires.free;
  end;

end;

function TOF_BTTACHEMODELE_GEN.CreationTaches: boolean;
var
  i, j, NumTache: integer;
  vAFOAssistants: TAFO_Ressources;
  vTOBArticles, vTOBAffaires, vTobMemo: TOB;
  QQ: tquery;
begin
  Result := false;
  fTOBTaches := TOB.Create('Les Taches', nil, -1);
  vAFOAssistants := TAFO_Ressources.Create;
  vTOBArticles := TOB.Create('Les Articles', nil, -1);
  vTOBAffaires := TOB.Create('Les Affaires', nil, -1);

  InitMoveProgressForm(nil, Ecran.Caption, '', fTOBAffaires.detail.count * fTOBModeles.detail.count, TRUE, TRUE);
  try
    // Création des tâches
    for i := 0 to fTOBAffaires.detail.count - 1 do
    begin
      NumTache := GetNumTache(fTOBAffaires.detail[i].getvalue('AFF_AFFAIRE'));
      for j := 0 to fTOBModeles.detail.count - 1 do
      begin
        CreationUneTache(fTOBAffaires.detail[i], fTOBModeles.detail[j], NumTache + j);
        if not MoveCurProgressForm(TexteMessage[5]) then
          Exit;
      end;
    end;
    // Valorisation des tâches
    if not MoveCurProgressForm(TexteMessage[7]) or (fTOBTaches.detail.count = 0) then
      Exit;
    for i := 0 to fTOBTaches.detail.count - 1 do
    begin
      ValorisationPlanning(fTOBTaches.detail[i], 'ATA', vAFOAssistants, vTOBAffaires, vTobArticles, true, true);
      //mcd 10/11/2005 on récupère le mémo si exite
      if fTobTaches.detail[i].getvalue('ATA_MODELETACHE') <> '' then
      begin
        QQ := OpenSql('SELECT * FROM LIENSOLE where LO_IDENTIFIANT="' + fTobTaches.detail[i].getvalue('ATA_MODELETACHE') + '"', true);
        if not QQ.EOF then
        begin
          vTobMemo := TOB.Create('les mémos', vTobMemo, -1);
          vTOBMemo.LoadDetailDB('LIENSOLE', '', '', QQ, False, False);
          for j := 0 to vtobMemo.detail.count - 1 do
          begin
            vtobMemo.detail[j].putvalue('LO_IDENTIFIANT', fTobTaches.detail[i].getvalue('ATA_AFFAIRE') + '/' + IntToSTr(fTobTaches.detail[i].getvalue('ATA_NUMEROTACHE')));
            vtobMemo.detail[j].putvalue('LO_TABLEBLOB', 'ATA');
          end;
        end;
        Ferme(QQ);
      end;
      if not MoveCurProgressForm(TexteMessage[7]) then
        Exit;
    end;
    // Enregistrement des tâches
    if not MoveCurProgressForm(TexteMessage[6]) then
      Exit;
    if VtobMemo <> nil then
    begin
      vTobMemo.setAllmodifie(true); //mcd  pour OK blob
      vTOBMemo.InsertDB(nil, true);
      FreeAndNil(vTobMemo);
    end;
    Result := fTOBTaches.InsertDB(nil, true);

  finally
    fTOBTaches.free;
    vAFOAssistants.free;
    vTOBArticles.free;
    vTOBAffaires.free;
    FiniMoveProgressForm;
  end;
end;

procedure TOF_BTTACHEMODELE_GEN.CreationUneTache(pTOBAffaire, pTOBModele: TOB; pNumTache: integer);
var
  vTobTache, VtobTacheJour, Tobdet: TOB;
  ii, indice: integer;
  Qq: tquery;
begin
  vTobTache := TOB.Create('TACHE', fTOBTaches, -1);
  vTobTacheJour := Tob.create('AFTACHEJOUR', nil, -1);
  Vtobtachejour.dupliquer(FTobJourModele, true, true);
  TOBCopieChamp(pTOBModele, vTobTache);
  //mcd 29/11/2005 12583
  QQ := Opensql('SELECT GA_ACTIVITEREPRISE FROM ARTICLE where GA_CODEARTICLE="' + vTobTAche.GetValue('ATA_CODEARTICLE') + '"', true);
  if not QQ.eof then
    vTobTAche.putValue('ATA_ACTIVITEREPRIS', QQ.Fields[0].AsString);
  ferme(QQ);
  //fin mcd 29/11/2005
  indice := 0;
  QQ := OpenSQL('SELECT MAX(ATJ_JOURNUMERO) FROM AFTACHEJOUR', true);
  if not QQ.Eof then
    Indice := QQ.Fields[0].AsInteger + 1;
  Ferme(QQ);
  for ii := 0 to vtobTachejour.detail.count - 1 do
  begin
    tobdet := vtobTacheJour.detail[ii];
    if vtobTache.getvalue('ATA_MODELETACHE') = tobdet.getvalue('ATJ_MODELETACHE') then
    begin // si multi sélection, il ne faut pas mettre tous les jour ssur totues les taches
      tobdet.putvalue('ATJ_AFFAIRE', pTOBAffaire.getvalue('AFF_AFFAIRE'));
      tobdet.putvalue('ATJ_NUMEROTACHE', pNumTache);
      tobdet.putvalue('ATJ_JOURNUMERO', Indice);
      tobdet.putvalue('ATJ_MODELETACHE', '');
      tobdet.putvalue('ATJ_TYPEJOUR', 'AFF');
      Inc(indice);
    end;
  end;
  vTobTacheJour.InsertorUpdateDb(False);
  vTobTacheJour.free;
  vTobTache.PutValue('ATA_TIERS', pTOBAffaire.getvalue('AFF_TIERS'));
  vTobTache.PutValue('ATA_AFFAIRE', pTOBAffaire.getvalue('AFF_AFFAIRE'));
  vTobTache.PutValue('ATA_AFFAIRE0', pTOBAffaire.getvalue('AFF_AFFAIRE0'));
  vTobTache.PutValue('ATA_AFFAIRE1', pTOBAffaire.getvalue('AFF_AFFAIRE1'));
  vTobTache.PutValue('ATA_AFFAIRE2', pTOBAffaire.getvalue('AFF_AFFAIRE2'));
  vTobTache.PutValue('ATA_AFFAIRE3', pTOBAffaire.getvalue('AFF_AFFAIRE3'));
  vTobTache.PutValue('ATA_AVENANT', pTOBAffaire.getvalue('AFF_AVENANT'));
  vTobTache.PutValue('ATA_DEVISE', pTOBAffaire.getvalue('AFF_DEVISE'));

  vTobTache.PutValue('ATA_NUMEROTACHE', pNumTache);
  if (vTobTache.getvalue('ATA_DATEDEBPERIOD') < pTOBAffaire.getvalue('AFF_DATEDEBUT'))
    or (vTobTache.getvalue('ATA_DATEDEBPERIOD') > pTOBAffaire.getvalue('AFF_DATEFIN'))
    then
    vTobTache.putvalue('ATA_DATEDEBPERIOD', pTOBAffaire.getvalue('AFF_DATEDEBUT'));
  if (vTobTache.getvalue('ATA_DATEFINPERIOD') > pTOBAffaire.getvalue('AFF_DATEFIN'))
    or (vTobTache.getvalue('ATA_DATEFINPERIOD') < pTOBAffaire.getvalue('AFF_DATEDEBUT'))
    then
    vTobTache.putvalue('ATA_DATEFINPERIOD', pTOBAffaire.getvalue('AFF_DATEFIN'));
  vTobTache.UpdateFieldsAgl;
  vTobTache.PutValue('ATA_DATECREATION', now);
end;

procedure TOF_BTTACHEMODELE_GEN.TOBCopieChamp(FromTOB, ToTOB: TOB);
var
  i_pos, i_ind1: integer;
  FieldNameTo, FieldNameFrom, St: string;
  PrefixeTo, PrefixeFrom: string;
begin
  PrefixeFrom := TableToPrefixe(FromTOB.NomTable);
  PrefixeTo := TableToPrefixe(ToTOB.NomTable);
  for i_ind1 := 1 to FromTOB.NbChamps do
  begin
    FieldNameFrom := FromTOB.GetNomChamp(i_ind1);
    St := FieldNameFrom;
    i_pos := Pos('_', St);
    System.Delete(St, 1, i_pos - 1);
    FieldNameTo := PrefixeTo + St;
    if ToTOB.FieldExists(FieldNameTo) then
      ToTOB.PutValue(FieldNameTo, FromTOB.GetValue(FieldNameFrom))
  end;
end;

{*************************************************************************
Auteur  ...... : AB
Créé le ...... : 04/03/2003
Modifié le ... :   /  /
Description ...: Recherche des modèles de tâches depuis les critères des MUL
Mots clefs ... : Fiche AFTACHEMODELE_MUL avec argument RECH
***************************************************************************}

function AFGetTacheModele(pEDTacheModele: THCritMaskEdit; pLequel: string): string;
var
  vEDTacheModele: THCritMaskEdit;
begin
  Result := '';
  if pEDTacheModele = nil then
  begin
    vEDTacheModele := THCritMaskEdit.create(nil);
    vEDTacheModele.Text := '';
  end
  else
    vEDTacheModele := pEDTacheModele;

  if getparamsocSecur('SO_BTAVANCEMODELE',false) or (pEDTacheModele = nil) then
    vEDTacheModele.Text := AGLLanceFiche('AFF', 'AFTACHEMODELE_MUL', pLequel, '', 'RECH')
  else
    LookUpCombo(vEDTacheModele);

  Result := vEDTacheModele.Text;
  if pEDTacheModele = nil then
    vEDTacheModele.Free;
end;

function AGLAFGetTacheModele(parms: array of variant; nb: integer): variant;
var
  CC: THCritMaskEdit;
  F: TForm;
begin
  F := TForm(Longint(Parms[0]));
  CC := THCritMaskEdit(F.FindComponent(string(Parms[1])));
  Result := AFGetTacheModele(CC, string(Parms[2]));
end;

initialization
  registerclasses([TOF_BTTACHEMODELE]);
  registerclasses([TOF_BTTACHEMODELE_GEN]);
  RegisterAglFunc('AFGetTacheModele', TRUE, 2, AGLAFGetTacheModele);
end.
