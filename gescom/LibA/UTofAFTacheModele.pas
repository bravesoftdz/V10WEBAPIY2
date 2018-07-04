{***********UNITE*************************************************
Auteur  ...... : AB
Créé le ...... : 19/02/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFTACHEMODELE_MUL ()
Mots clefs ... : TOF;AFTACHEMODELE
*****************************************************************}
Unit UtofAfTacheModele ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFDEF EAGLCLIENT}
      eMul,Maineagl,
{$ELSE}
      db,dbTables,mul,FE_Main,
{$ENDIF}
     forms, 
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,HTB97,ed_tools,M3FP,
     UTOF,UTOB,LookUp,paramsoc,
     GereTobInterne,EntGC,
     AfUtilArticle,UtilMulTrt,UtilTaches,UtilGC,
     UTofAfBaseCodeAffaire,DicoAf,ConfidentAffaire,UAFO_Ressource ;

Type
  TOF_AFTACHEMODELE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    private
      bMulti : Boolean;
      fTOBModeles : TOB;
      procedure BOuvrirOnClick (Sender: TObject);
  end ;

  TOF_AFTACHEMODELE_GEN = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    private
      bMultiModele : Boolean;    
      fTOBTaches, fTOBAffaires, fTOBModeles : TOB;
      procedure SelectionModeleTache (Sender: TObject);
      procedure BOuvrirOnClick(Sender: TObject);
      procedure MultiModeleClick (Sender: TObject);      
      function  CreationTaches : boolean;
      procedure CreationUneTache(pTOBAffaire,pTOBModele:TOB;pNumTache:integer);
      procedure TOBCopieChamp(FromTOB, ToTOB : TOB);
  end ;

procedure AFLanceFicheAFTacheModele_Mul;
procedure AFLanceFicheAFTacheModele_Gen;
Function AFGetTacheModele (pEDTacheModele: THCritMaskEdit; pLequel : string) : string ;

const
// libellés des messages
TexteMessage: array[1..7] of string 	= (
    {1}  'Vous n''avez pas sélectionné les modèles de tâche.'
    {2} ,'Vous avez sélectionné %d tâches à créer dans %d affaires, voulez-vous continuer ?'
    {3} ,'La création des tâches a été effectuée avec succès.'
    {4} ,'La création des tâches ne s''est pas effectuée.'
    {5} ,'Création des tâches en cours...'
    {6} ,'Enregistrement des tâches en cours...'
    {7} ,'Valorisation des tâches en cours...'
    );

Implementation

procedure AFLanceFicheAFTacheModele_Mul;
begin
  AGLLanceFiche('AFF','AFTACHEMODELE_MUL','','','');
end;

procedure AFLanceFicheAFTacheModele_Gen;
begin
  AGLLanceFiche('AFF','AFTACHEGEN_MUL','','','');
end;

procedure TOF_AFTACHEMODELE.OnArgument (S : String ) ;
begin
  Inherited ;
  if S='MULTI' then
  begin
    bMulti := true;
  {$IFDEF EAGLCLIENT}
    SetControlProperty('Fliste','MultiSelect',true);
  {$ELSE}
    SetControlProperty('Fliste','Multiselection',true);
  {$ENDIF}
    setcontrolVisible('BSELECTALL', true);
    setcontrolVisible('BINSERT', false);
    setcontrolText('MODE','MULTI');
    Ecran.Caption := TraduitGA('Sélection des modèles de tâches');
    TFMul(Ecran).BOuvrir.OnClick:= BOuvrirOnClick;
  end
  else if S='RECH' then
  begin
    setcontrolVisible('BSELECTALL', false);
    setcontrolVisible('BINSERT', false);
    setcontrolText('MODE','RECH');
    Ecran.Caption := TraduitGA('Recherche des modèles de tâches');
  end;
  {$IFDEF CCS3}
  if (getcontrol('PTEXTELIBRE') <> Nil) then SetControlVisible ('PTEXTELIBRE', False);
  if (getcontrol('PZONE') <> Nil)  then SetControlVisible ('PZONE', False);
  {$ENDIF}

end ;

procedure TOF_AFTACHEMODELE.OnLoad ;
begin
  Inherited ;
  SetControlProperty('AFM_TYPEARTICLE','Plus',PlusTypeArticle) ;
  // traduction champs libres
  GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'AFM_LIBRETACHE', 10, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFM_DATELIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFM_CHARLIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFM_VALLIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'AFM_BOOLLIBRE', 3, '_');
end ;

procedure TOF_AFTACHEMODELE.BOuvrirOnClick(Sender: TObject);
begin
  fTOBModeles := MaTobInterne ('Liste ModeleTaches');
  fTOBModeles.cleardetail;
  TraiteEnregMulTable (TFMul(Ecran), 'SELECT * FROM AFMODELETACHE', 'AFM_MODELETACHE', 'AFMODELETACHE',
                          'AFM_MODELETACHE','AFMODELETACHE', fTOBModeles, True);
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

procedure TOF_AFTACHEMODELE_GEN.NomsChampsAffaire(var Aff, Aff0, Aff1,
  Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
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

procedure TOF_AFTACHEMODELE_GEN.OnArgument(S: String);
begin
  inherited;
  if Not(ModifAffaireAutorise) then SetControlText('XXAction','ACTION=CONSULTATION');
  fTOBModeles := MaTobInterne ('Liste ModeleTaches');
  TFMul(Ecran).BOuvrir.OnClick:= BOuvrirOnClick;
  TToolBarButton97(GetControl('BSELECTMOD')).onClick := SelectionModeleTache;
  TCheckBox(GetControl('MULTIMODELE')).onClick := MultiModeleClick;
  {$IFDEF CCS3}
  if (getcontrol('PSTAT') <> Nil) then SetControlVisible ('PSTAT', False);
  if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
  {$ENDIF}
end;

procedure TOF_AFTACHEMODELE_GEN.OnClose;
begin
  inherited;
  if (VH_GC.AFTobInterne <> nil) then
    DetruitMaTobInterne ('Liste ModeleTaches');
end;

procedure TOF_AFTACHEMODELE_GEN.SelectionModeleTache(Sender: TObject);
var CBModele :THValcomboBox;
    i : integer;
begin
  AGLLanceFiche('AFF','AFTACHEMODELE_MUL','','','MULTI');
  CBModele := THValcomboBox(GetControl('MODELETACHEMULTI'));
  CBModele.Items.clear;
  CBModele.Values.clear;
  CBModele.ItemIndex := -1;
  if (fTOBModeles <> nil) and (fTOBModeles.detail.Count > 0) then
  begin
    for i:=0 to fTOBModeles.detail.count-1 do
    begin
      CBModele.Items.add (fTOBModeles.detail[i].GetValue('AFM_LIBELLETACHE1'));
      CBModele.Values.add (fTOBModeles.detail[i].GetValue('AFM_MODELETACHE'));
    end;
    CBModele.ItemIndex := 0;
  end;
end;

procedure TOF_AFTACHEMODELE_GEN.MultiModeleClick (Sender: TObject);
begin
  bMultiModele := (GetControlText('MULTIMODELE')='X');
  SetControlVisible('BSELECTMOD',(GetControlText('MULTIMODELE')='X') );
  SetControlVisible('MODELETACHE',(GetControlText('MULTIMODELE')='-') );
  SetControlVisible('MODELETACHEMULTI',(GetControlText('MULTIMODELE')='X') );
end;

procedure TOF_AFTACHEMODELE_GEN.BOuvrirOnClick(Sender: TObject);
var vSelect : string;
    TobDetModele : Tob;
begin
  try
    SourisSablier;
    fTOBAffaires := Tob.create ('Liste Affaires',nil,-1);
    vSelect := 'SELECT * FROM AFFAIRE';
    TraiteEnregMulTable (TFMul(Ecran), vSelect, 'AFF_AFFAIRE', 'AFFAIRE', 'AFF_AFFAIRE', 'AFFAIRE', fTOBAffaires, True);
    if fTOBAffaires.detail.count = 0 then Exit;
    if not bMultiModele then
    begin
      TobDetModele := TOB.Create ('AFMODELETACHE', fTOBModeles,-1);
      TobDetModele.InitValeurs;
      TobDetModele.SelectDB ('"' +GetControltext('MODELETACHE')+ '"',nil );
    end
    else
    if fTOBModeles.detail.count = 0 then
    begin
      PgiInfoAf(TexteMessage[1], Ecran.Caption);
      Exit;
    end;

    if (PGIAskAF (format(TexteMessage[2],[fTOBModeles.detail.count,fTOBAffaires.detail.count]),Ecran.Caption) <> mrYes) then Exit;
    if CreationTaches then
      PgiInfoAf(TexteMessage[3], Ecran.Caption)
    else PgiInfoAf(TexteMessage[4], Ecran.Caption);

  finally
    SourisNormale;
    fTOBAffaires.free;
  end;
end;

function TOF_AFTACHEMODELE_GEN.CreationTaches : boolean;
var i,j,NumTache : integer;
    vAFOAssistants : TAFO_Ressources;
    vTOBArticles,vTOBAffaires :TOB;
begin
  Result := false;
  fTOBTaches := TOB.Create('Les Taches',Nil,-1);
  vAFOAssistants  := TAFO_Ressources.Create;
  vTOBArticles := TOB.Create('Les Articles',Nil,-1);
  vTOBAffaires := TOB.Create('Les Affaires',Nil,-1);

  InitMoveProgressForm (nil,Ecran.Caption,'',fTOBAffaires.detail.count * fTOBModeles.detail.count,TRUE,TRUE) ;
  try
    // Création des tâches
    for i := 0 to fTOBAffaires.detail.count-1 do
    begin
      NumTache := GetNumTache (fTOBAffaires.detail[i].getvalue ('AFF_AFFAIRE'));
      for j := 0 to fTOBModeles.detail.count-1 do
      begin
        CreationUneTache (fTOBAffaires.detail[i], fTOBModeles.detail[j], NumTache+j);
        if Not MoveCurProgressForm(TexteMessage[5]) then Exit ;
      end;
    end;
    // Valorisation des tâches
    if Not MoveCurProgressForm(TexteMessage[7]) or (fTOBTaches.detail.count = 0) then Exit ;
    for i := 0 to fTOBTaches.detail.count-1 do
    begin
      Valorisation(fTOBTaches.detail[i], 'ATA', vAFOAssistants, vTOBAffaires, vTobArticles);
      if Not MoveCurProgressForm(TexteMessage[7]) then Exit ;
    end;
    // Enregistrement des tâches
    if Not MoveCurProgressForm(TexteMessage[6]) then Exit ;
    Result := fTOBTaches.InsertDB(nil,true);

  finally
    fTOBTaches.free;
    vAFOAssistants.free;
    vTOBArticles.free;
    vTOBAffaires.free;
    FiniMoveProgressForm ;
  end;
end;

procedure TOF_AFTACHEMODELE_GEN.CreationUneTache(pTOBAffaire, pTOBModele : TOB;pNumTache : integer);
var vTobTache :TOB;
begin
  vTobTache := TOB.Create('TACHE', fTOBTaches, -1);
  TOBCopieChamp (pTOBModele, vTobTache);
  vTobTache.PutValue ('ATA_TIERS',       pTOBAffaire.getvalue ('AFF_TIERS'));
  vTobTache.PutValue ('ATA_AFFAIRE',     pTOBAffaire.getvalue ('AFF_AFFAIRE'));
  vTobTache.PutValue ('ATA_AFFAIRE0',    pTOBAffaire.getvalue ('AFF_AFFAIRE0'));
  vTobTache.PutValue ('ATA_AFFAIRE1',    pTOBAffaire.getvalue ('AFF_AFFAIRE1'));
  vTobTache.PutValue ('ATA_AFFAIRE2',    pTOBAffaire.getvalue ('AFF_AFFAIRE2'));
  vTobTache.PutValue ('ATA_AFFAIRE3',    pTOBAffaire.getvalue ('AFF_AFFAIRE3'));
  vTobTache.PutValue ('ATA_AVENANT',     pTOBAffaire.getvalue ('AFF_AVENANT'));
  vTobTache.PutValue ('ATA_DEVISE',      pTOBAffaire.getvalue ('AFF_DEVISE'));

  vTobTache.PutValue ('ATA_NUMEROTACHE', pNumTache);
  if (vTobTache.getvalue('ATA_DATEDEBPERIOD') < pTOBAffaire.getvalue('AFF_DATEDEBUT'))
  or (vTobTache.getvalue('ATA_DATEDEBPERIOD') > pTOBAffaire.getvalue('AFF_DATEFIN'))
  then
    vTobTache.putvalue('ATA_DATEDEBPERIOD',pTOBAffaire.getvalue('AFF_DATEDEBUT'));
  if (vTobTache.getvalue('ATA_DATEFINPERIOD') > pTOBAffaire.getvalue('AFF_DATEFIN'))
  or (vTobTache.getvalue('ATA_DATEFINPERIOD') < pTOBAffaire.getvalue('AFF_DATEDEBUT'))
  then
    vTobTache.putvalue('ATA_DATEFINPERIOD',pTOBAffaire.getvalue('AFF_DATEFIN'));
  // ajout des parametres de decalage 
  vTobTache.PutValue('ATA_NBJOURSDECAL', GetParamSoc('SO_AFJOURSDECAL'));
  vTobTache.PutValue('ATA_METHODEDECAL', GetParamSoc('SO_AFJOURSPLANIF'));
  vTobTache.UpdateFieldsAgl;
  vTobTache.PutValue('ATA_DATECREATION',now);
end;

procedure TOF_AFTACHEMODELE_GEN.TOBCopieChamp(FromTOB, ToTOB : TOB);
var i_pos,i_ind1: integer;
    FieldNameTo,FieldNameFrom,St:string;
    PrefixeTo,PrefixeFrom : string;
begin
  PrefixeFrom := TableToPrefixe (FromTOB.NomTable);
  PrefixeTo := TableToPrefixe (ToTOB.NomTable);
  for i_ind1 := 1 to FromTOB.NbChamps do
  begin
    FieldNameFrom := FromTOB.GetNomChamp(i_ind1);
    St := FieldNameFrom ;
    i_pos := Pos ('_', St) ;
    System.Delete (St, 1, i_pos-1) ;
    FieldNameTo := PrefixeTo + St ;
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

Function AFGetTacheModele (pEDTacheModele: THCritMaskEdit; pLequel : string) : string ;
var vEDTacheModele : THCritMaskEdit;
begin
  Result := '' ;
  if pEDTacheModele = nil then
  begin
    vEDTacheModele := THCritMaskEdit.create (nil);
    vEDTacheModele.Text := '';
  end else vEDTacheModele := pEDTacheModele;

  if getparamsoc('SO_AFAVANCEMODELE') or (pEDTacheModele = nil) then
       vEDTacheModele.Text := AGLLanceFiche ('AFF', 'AFTACHEMODELE_MUL', pLequel, '', 'RECH')
  else LookUpCombo (vEDTacheModele);

  Result := vEDTacheModele.Text;
  if pEDTacheModele = nil then vEDTacheModele.Free;
end ;

Function AGLAFGetTacheModele (parms: array of variant; nb: integer ) : variant ;
var CC : THCritMaskEdit;
    F : TForm;
begin
F := TForm (Longint (Parms [0]));
CC := THCritMaskEdit (F.FindComponent (string (Parms [1])));
Result := AFGetTacheModele (CC, string (Parms [2]));
end;

Initialization
  registerclasses ( [ TOF_AFTACHEMODELE ] ) ;
  registerclasses ( [ TOF_AFTACHEMODELE_GEN ] ) ;
  RegisterAglFunc ( 'AFGetTacheModele', TRUE, 2, AGLAFGetTacheModele);
end.
