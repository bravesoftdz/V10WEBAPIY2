{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/01/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BPIECEDEMPRIX_MUL ()
Mots clefs ... : TOF;BPIECEDEMPRIX_MUL
*****************************************************************}
Unit BPIECEDEMPRIX_MUL_TOF;

Interface

Uses StdCtrls,
     Controls,
     Classes,

{$IFDEF EAGLCLIENT}
     MaineAGL,
     eMul,
     uTob,
{$ELSE}
     DBCtrls, Db,
     {$IFNDEF DBXPRESS}
     dbTables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF DBXPRESS}
     fe_main,
     HDB,
     Mul,
{$ENDIF EAGLCLIENT}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     HTB97,
     uTOFComm,
     uEntCommun,
     uTob,
     Menus,
     Udefexport,
     UTofAfBasePiece_Mul,
     UTOF;

Type
  TtypeModeTrait = (TtmtSelect,TtmtModif);
  TOF_BPIECEDEMPRIX_MUL = Class (TOF_AFBASEPIECE_MUL)
    procedure OnNew                    ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDelete                 ; Override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  Private
    //
    fTtypeTrait : TtypeModeTrait;
   	fTypeExport : TTypeExport;
    fModeExport : TmodeExport;
    //
    BtCreate    : TToolbarButton97;
    BtDelete    : TToolBarButton97;
    BtExportXLS : TToolBarButton97;
    BTSaisieDde : TToolBarButton97;
    //
    Grille      : THDBGrid;
    //
    NaturePieceg: THValComboBox;
    Souche      : THEdit;
    Numero      : THEdit;
    Indiceg     : THEdit;
    //
    TheAction   : String;
    //
    Unique      : Integer;
    Cledoc      : R_CLEDOC;
    LibDdePrix  : String;
    //
    procedure MnExportOnClick(Sender: Tobject);
    procedure OnDblClick(Sender: Tobject);
    procedure SuppressionDdePrixDirecte(Unique: Integer);
    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Critere: String);
    procedure BtSaisieDdeOnClick(Sender: Tobject);

  public
    //
    procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
    //
  end ;

const
  // libellés des messages de la TOM Affaire
  TexteMessage : array [1..6] of string = (
    {1} 'Confirmez-vous la suppression de cette demande de prix ?',
    {2} 'Erreur de suppression de la demandes de Prix',
    {3} 'Type d''export non défini',
    {4} 'Veuillez renseigner l''emplacement de stockage des exports',
    {5} 'Modèle d''export non défini',
    {6} 'Excel n''est pas installé sur ce poste'
    ) ;

Implementation
uses  UtilTOBPiece,
      UDemandePrix,
      CalcOLEGenericBTP,
      Paramsoc,
      TiersUtil,
      UfactExportXLS, Math;

procedure TOF_BPIECEDEMPRIX_MUL.OnNew ;
begin
  Inherited ;

  CreateDemandePrix(Cledoc, Unique, 'CREATION');

  RefreshDB;

end ;


procedure TOF_BPIECEDEMPRIX_MUL.OnDelete;
begin
  inherited;

  Unique := Grille.Datasource.Dataset.FindField('BPP_UNIQUE').AsInteger;

  //lecture entete demande de prix
  if (PGIAsk(TexteMessage[1], ecran.Caption) = mrYes) then
  begin
    SuppressionDdePrixDirecte(Unique);
    RefreshDB;
  end;

end ;

Procedure TOF_BPIECEDEMPRIX_MUL.SuppressionDdePrixDirecte(Unique : Integer);
begin

  BEGINTRANS;
  TRY
    SupprimeDdePrix(cledoc, Unique,'PIECEDEMPRIX',    'BPP', TTdPieceDemPrix);
    SupprimeDdePrix(Cledoc, Unique,'ARTICLEDEMPRIX',  'BDP', TTdArticleDemPrix);
    SupprimeDdePrix(Cledoc, Unique,'DETAILDEMPRIX',   'BD0', TTdDetailDemPrix);
    SupprimeDdePrix(Cledoc, Unique,'FOURLIGDEMPRIX',  'BD1', TTdFournDemprix);
    COMMITTRANS;
  EXCEPT
    ROLLBACK;
    PGIInfo(TexteMessage[2], Ecran.caption);
  END;

end;

procedure TOF_BPIECEDEMPRIX_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnArgument (S : String ) ;
Var Critere : String;
    Champ   : String;
    Valeur  : Variant;
    X       : Integer;
    CC      : THValComboBox;
begin

  fMulDeTraitement := true;

  Inherited ;

  fTtypeTrait := TtmtModif;
  if pos('SELECTION',S)> 0 then fTtypeTrait := TtmtSelect;

	fTypeExport := TteUndefined;
  fModeExport := TmeUndefined;

  if assigned(Getcontrol('BInsert')) then
  begin
    BtCreate := TToolbarButton97(ecran.FindComponent('BInsert'));
    BtCreate.Visible := False;
  end;

  if Assigned(GetControl('BDelete')) then
  begin
    BtDelete := TToolbarButton97(ecran.findComponent('BDelete'));
    BtDelete.Enabled := True;
    BtDelete.Visible := True;
  end;

  if Assigned(GetControl('Fliste')) then
  begin
    Grille := THDBGrid(ecran.FindComponent('Fliste'));
    Grille.OnDblClick := OnDblClick;
  end;

  if assigned(GetControl('BEXPORTS')) then
  begin
    BtExportXLS := TToolbarButton97(Ecran.Findcomponent('BEXPORTS'));
    BtExportXLS.OnClick := MnExportOnClick;
  end;

  if assigned(GetControl('BTSAISIEDDE')) then
  begin
    BtSaisieDde := TToolbarButton97(Ecran.Findcomponent('BTSAISIEDDE'));
    BtSaisieDde.OnClick := BtSaisieDdeOnClick;
  end;

  NaturePieceg := THValComboBox(Ecran.FindComponent('GP_NATUREPIECEG'));
  Souche       := THEdit(Ecran.FindComponent('GP_SOUCHE'));
  IndiceG      := THEdit(Ecran.FindComponent('GP_INDICEG'));
  Numero       := THEdit(Ecran.FindComponent('GP_NUMERO'));

  //Récupération valeur de argument
  Critere:=(Trim(ReadTokenSt(S)));

  while (Critere <> '') do
  begin
    if Critere <> '' then
    begin
      X := pos (';', Critere) ;
      if x = 0 then
        X := pos ('=', Critere) ;
      if x <> 0 then
      begin
        Champ := copy (Critere, 1, X - 1) ;
        Valeur:= Copy (Critere, X + 1, length (Critere) - X) ;
        ControleChamp(champ, valeur);
      end
    end;
    ControleCritere(Critere);
    Critere   := (Trim(ReadTokenSt(S)));
  end;

  //Gestion Restriction Domaine et Etablissements
  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;

end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BPIECEDEMPRIX_MUL.OnDblClick(Sender: Tobject);
begin

  if (fTtypeTrait = TtmtModif ) then
  begin
    Cledoc.NaturePiece := Grille.Datasource.Dataset.FindField('GP_NATUREPIECEG').AsString;
    Cledoc.Souche      := Grille.Datasource.Dataset.FindField('GP_SOUCHE').AsString;
    Cledoc.NumeroPiece := Grille.Datasource.Dataset.FindField('GP_NUMERO').AsInteger;
    Cledoc.Indice      := Grille.Datasource.Dataset.FindField('GP_INDICEG').AsInteger;
    Unique             := Grille.Datasource.Dataset.FindField('BPP_UNIQUE').AsInteger;
    //
    CreateDemandePrix(Cledoc, Unique, 'MODIFICATION');
    //
    RefreshDB();
  end
  else
  begin
    // selection et sortie
    Tfmul(ecran).Retour := IntToStr(Grille.Datasource.Dataset.FindField('BPP_UNIQUE').AsInteger);
    ecran.Close;
  end;
end;

Procedure TOF_BPIECEDEMPRIX_MUL.BtSaisieDdeOnClick(Sender : Tobject);
var Argument : String;
Begin

  Argument := 'NATUREPIECEG='       + Grille.Datasource.Dataset.FindField('GP_NATUREPIECEG').AsString;
  Argument := Argument + ';SOUCHE=' + Grille.Datasource.Dataset.FindField('GP_SOUCHE').AsString;
  Argument := Argument + ';NUMERO=' + IntToStr(Grille.Datasource.Dataset.FindField('GP_NUMERO').AsInteger);
  Argument := Argument + ';INDICEG='+ IntToStr(Grille.Datasource.Dataset.FindField('GP_INDICEG').AsInteger);
  Argument := Argument + ';UNIQUE=' + IntToStr(Grille.Datasource.Dataset.FindField('BPP_UNIQUE').AsInteger);

  AGLLanceFiche('BTP','BTVALIDEDDEPRIX','','',Argument);

  refreshDB();

end;

Procedure TOF_BPIECEDEMPRIX_MUL.MnExportOnClick(Sender : Tobject);
Var unique  : Integer;
    CleDoc  : R_CLedoc;
begin

  //génération de la tob nécessaire à l'envoi dans excel
  fillchar(cledoc,sizeof(cledoc),0);
  //
  cledoc.NaturePiece  := Grille.Datasource.Dataset.FindField('GP_NATUREPIECEG').AsString;
  cledoc.souche       := Grille.Datasource.Dataset.FindField('GP_SOUCHE').AsString;
  cledoc.NumeroPiece  := Grille.Datasource.Dataset.FindField('GP_NUMERO').AsInteger;
  cledoc.Indice       := Grille.Datasource.Dataset.FindField('GP_INDICEG').AsInteger;

  Unique      := Grille.Datasource.Dataset.FindField('BPP_UNIQUE').AsInteger;
  LibDdePrix  := Grille.Datasource.Dataset.FindField('BPP_LIBELLE').AsString;

  ExporteDocumentToExcel(Cledoc, Unique, LibDdePrix);

  RefreshDB();

end;

procedure TOF_BPIECEDEMPRIX_MUL.ControleChamp(Champ, Valeur: String);
begin

  if Champ = 'BPP_NATUREPIECEG' then
    NaturePieceg.Value := Valeur
  else if Champ = 'BPP_SOUCHE' then
    Souche.text       := Valeur
  else if Champ = 'BPP_NUMERO' then
    Numero.Text := Valeur
  else if Champ = 'BPP_INDICEG' then
    Indiceg.text      := Valeur
  else if Champ = 'BPP_UNIQUE' then
    Unique            := StrToInt(Valeur)
  else if champ = 'ACTION' then
    TheAction         := Valeur
  Else if champ = 'STATUT' then
  begin
    if Valeur = 'APP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'W');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="W"');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
      SetControlText('TGPAFFAIRE', 'Appel');
    end
    Else if valeur = 'INT' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'I');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="I"');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
      SetControlText('TGPAFFAIRE', 'Contrat');
    end
    Else if valeur = 'AFF' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'A')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'A');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A", "")');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
      SetControlText('TGPAFFAIRE', 'Chantier');
    end
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControltext('XX_WHERE',' AND AFF_AFFAIRE0 IN ("A", "W")')
      else if assigned(GetControl('AFFAIRE0')) then SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) in ("A", "W")');
    end
    Else if valeur = 'PRO' then
    Begin
      //
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'P');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="P"');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
      SetControlText('TGPAFFAIRE', 'Appel d''offre');
    end
    else
    Begin
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlText('TGPAFFAIRE', 'Affaire');
    end
  end;
end;

procedure TOF_BPIECEDEMPRIX_MUL.ControleCritere(Critere: String);
begin
end;

Procedure TOF_BPIECEDEMPRIX_MUL.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
  inherited;

  Aff :=THEdit(GetControl('GP_AFFAIRE'));
  Aff0:=THEdit(GetControl('AFFAIRE0')) ;
  Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
  Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
  Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
  Aff4:=THEdit(GetControl('GP_AVENANT'))  ;

END ;


Initialization
  registerclasses ( [ TOF_BPIECEDEMPRIX_MUL ] ) ;
end.

