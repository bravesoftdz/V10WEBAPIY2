unit Synchro;

interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
{$IFDEF EAGLCLIENT}
{$ELSE}
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
    Vierge,
    HSysMenu,
    HTB97,
    StdCtrls,
    Hctrls,
    HPanel,
    UiUtil,
    ImgList,
    hmsgbox,
    UTob,
    HEnt1, UTobView, HImgList;

procedure SynchroTaLi;

type
  TFSynchro = class(TFVierge)
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
		cbo_CPTA: THValComboBox;
    cbo_GCFournisseur: THValComboBox;
    cbo_GCClient: THValComboBox;
    btn_Synchro: TToolbarButton97;
    iml_Liste: THImageList;
    MsgBox: THMsgBox;
    TobViewer1: TTobViewer;
    rbt_Client: TRadioButton;
    rbt_Fournisseur: TRadioButton;
    procedure btn_SynchroClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure rbt_ClientClick(Sender: TObject);
    procedure rbt_FournisseurClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    cbo_GC: THValComboBox;
    szChoixCodValue : String;
    szChoixExtValue : String;
    szChampTiersGC : String;
    szNature : String;
    procedure RempliCombo;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FSynchro: TFSynchro;

implementation

{$R *.DFM}

{=======================================================================================}
Procedure SynchroTaLi ;
var
  FSynchro: TFSynchro;
  PP : THPanel;
begin
  FSynchro := TFSynchro.Create(Application);
  PP:=FindInsidePanel ;
  if PP=Nil then begin
    try
      FSynchro.ShowModal;
    finally
      FSynchro.Free;
    end;
    end
  else begin
    InitInside(FSynchro,PP) ;
    FSynchro.Show ;
  end;
end;

{=======================================================================================}
procedure TFSynchro.btn_SynchroClick(Sender: TObject);
begin
  inherited;
  if (btn_Synchro.ImageIndex = 0) then
		btn_Synchro.ImageIndex := 1
  else
  	btn_Synchro.ImageIndex := 0;
end;

{=======================================================================================}
procedure TFSynchro.BValiderClick(Sender: TObject);
var
  TobCommun : Tob;
  TobChoixCod : Tob;
  TobNatCpte : Tob;
  TobChoixExt : Tob;
  TobFille : Tob;
  szLibelle : String;
  szAbrege : String;
  i : Integer;
  iLongueur : Integer;
begin
  inherited;

  // Vérifie si les combos sont renseignés
  if (cbo_CPTA.Value = '') or (cbo_GC.Value = '') then begin
    if ((btn_Synchro.ImageIndex = 0) and (cbo_CPTA.Value = '')) or
       ((btn_Synchro.ImageIndex = 1) and (cbo_GC.Value = '')) then begin
      MsgBox.Execute(1,Caption,'');		// Vous devez sélectionner une table de destination.
      Exit;
      end
    else begin
      MsgBox.Execute(0,Caption,'');		// Vous devez sélectionner une table source.
      Exit;
    end;
  end;

  // Demande confirmation
  if (btn_Synchro.ImageIndex = 1) then begin
    // Si la table libre en comptabilité contient des informations auxquelles des tiers font reférence, ce traitement entrainera une perte d'informations. Désirez-vous continuer ?
    if (MsgBox.Execute(2,Caption,'') = mrNo) then Exit;
    end
  else begin
    // Si la table libre en gestion commerciale contient des informations auxquelles des tiers font reférence, ce traitement entrainera une perte d'informations. Désirez-vous continuer ?
    if (MsgBox.Execute(3,Caption,'') = mrNo) then Exit;
  end;
  iLongueur := 0;
  BValider.Enabled := False;
  SourisSablier;

  // Récupère les données de la table COMMUN (Compta) et de la table CHOIXCOD(GC)
  TobCommun := Tob.Create('COMMUN',nil,-1);
  TobCommun.LoadDetailDB('COMMUN','"NAT";"T0'+cbo_CPTA.Value+'"','',nil,False,False);

  TobChoixCod := Tob.Create('CHOIXCOD',nil,-1);	// CT: Client ; FT:Fournisseur
  TobChoixCod.LoadDetailDB('CHOIXCOD','"ZLT";"'+szChoixCodValue+cbo_GC.Value+'"','',nil,False,False);

  TobNatCpte := Tob.Create('NATCPTE',nil,-1);
  TobChoixExt := Tob.Create('CHOIXEXT',nil,-1);

  // De GC vers compta
  BeginTrans;
  if (btn_Synchro.ImageIndex = 1) then begin
    // Suppression des références à la table dans la table TIERS
    ExecuteSql('UPDATE TIERS SET T_TABLE'+cbo_CPTA.Value+'="" WHERE T_TABLE'+cbo_CPTA.Value+'<>""');

    // Suppression des données de la table libre destination
    ExecuteSql('DELETE NATCPTE WHERE NT_TYPECPTE = "T0'+cbo_CPTA.Value+'"');

    // STRUCTURE DE LA TABLE LIBRE
    if (TobChoixCod.Detail.Count > 0) then begin
      szLibelle := TobChoixCod.Detail[0].GetValue('CC_LIBELLE');
      szAbrege := TobChoixCod.Detail[0].GetValue('CC_ABREGE');
    end;

    // Crée l'enregistrement s'il n'existe pas
    if (TobCommun.Detail.Count = 0) then begin
      TobFille := Tob.Create('CHOIXCOD',TobCommun,-1);
      TobFille.PutValue('CC_TYPE','NAT');
      TobFille.PutValue('CC_CODE','T0'+cbo_CPTA.Value);
    end;

    if (szLibelle = '.-') then begin	// Table cachée
      TobCommun.Detail[0].PutValue('CC_LIBELLE','Table n° '+IntToStr(StrToInt(cbo_CPTA.Value)+1));
      TobCommun.Detail[0].PutValue('CC_ABREGE','-');
      end
    else begin
      TobCommun.Detail[0].PutValue('CC_LIBELLE',szLibelle);
      TobCommun.Detail[0].PutValue('CC_ABREGE','X');
    end;
    TobCommun.Detail[0].PutValue('CC_LIBRE','6');
    if (TobCommun.InsertOrUpdateDB(False) = False) then begin
      // Impossible de mettre à jour la structure de la table libre
      MsgBox.Execute(5,Caption,'');
      RollBack;
      TobCommun.Free;
      TobChoixCod.Free;
      TobNatCpte.Free;
      TobChoixExt.Free;
      BValider.Enabled := True;
      SourisNormale;
      Exit;
    end;

    // Charge les données de la table libre source LT: Client ; LC: Fournisseur
    TobChoixExt.LoadDetailDB('CHOIXEXT','"'+szChoixExtValue+cbo_GC.Value+'"','',nil,False,False);

    // Recopie les données dans la table destination
    for i := 0 to TobChoixExt.Detail.Count-1 do begin
      TobFille := Tob.Create('NATCPTE',TobNatCpte,-1);
      TobFille.PutValue('NT_TYPECPTE','T0'+cbo_CPTA.Value);
      TobFille.PutValue('NT_NATURE',TobChoixExt.Detail[i].GetValue('YX_CODE'));
      TobFille.PutValue('NT_LIBELLE',TobChoixExt.Detail[i].GetValue('YX_ABREGE'));
    end;
    if (TobNatCpte.InsertOrUpdateDB(False) = False) then begin
      // Impossible de mettre à jour les données de la table libre
      MsgBox.Execute(6,Caption,'');
      RollBack;
      TobCommun.Free;
      TobChoixCod.Free;
      TobNatCpte.Free;
      TobChoixExt.Free;
      BValider.Enabled := True;
      SourisNormale;
      Exit;
    end;

    // Recopie des données dans la table TIERS
    ExecuteSql('UPDATE TIERS SET T_TABLE'+cbo_CPTA.Value+'=(SELECT YTC_TABLELIBRE'+szChampTiersGC+cbo_GC.Value+' FROM TIERSCOMPL WHERE TIERSCOMPL.YTC_AUXILIAIRE = TIERS.T_AUXILIAIRE) WHERE TIERS.T_AUXILIAIRE IN (SELECT YTC_AUXILIAIRE FROM TIERSCOMPL WHERE YTC_TABLELIBRE' + szChampTiersGC + cbo_GC.Value + ' <> "")');
    end
  // De compta vers GC
  else begin
    // Suppresssion des références à la table dans la table TIERSCOMPL
    ExecuteSql('UPDATE TIERSCOMPL SET YTC_TABLELIBRETIERS'+cbo_GC.Value+'="" WHERE YTC_TABLELIBRETIERS'+cbo_GC.Value+'<>""');

    // Suppression des données de la table libre destination
    ExecuteSql('DELETE CHOIXEXT WHERE YX_TYPE = "'+szChoixExtValue+cbo_GC.Value+'"');

    // STRUCTURE DE LA TABLE LIBRE
    if (TobCommun.Detail.Count > 0) then begin
      szLibelle := TobCommun.Detail[0].GetValue('CC_LIBELLE');
      szAbrege := TobCommun.Detail[0].GetValue('CC_ABREGE');
      iLongueur := StrToInt(TobCommun.Detail[0].GetValue('CC_LIBRE'));
    end;

    // Crée l'enregistrement s'il n'existe pas
    if (TobChoixCod.Detail.Count = 0) then begin
      TobFille := Tob.Create('CHOIXCOD',TobChoixCod,-1);
      TobFille.PutValue('CC_TYPE','ZLT');
      TobFille.PutValue('CC_CODE',szChoixCodValue+cbo_GC.Value);
    end;

    if (szAbrege = '-') then begin	// Table cachée
      TobChoixCod.Detail[0].PutValue('CC_LIBELLE','.-');
      if (cbo_GC.Value <> 'A') then
      	TobChoixCod.Detail[0].PutValue('CC_ABREGE','Table libre '+IntToStr(StrToInt(cbo_GC.Value)+1))
      else
      	TobChoixCod.Detail[0].PutValue('CC_ABREGE','Table libre 10');
      end
    else begin
      TobChoixCod.Detail[0].PutValue('CC_LIBELLE',szLibelle);
      TobChoixCod.Detail[0].PutValue('CC_ABREGE',Copy(szLibelle,0,17));
    end;
    TobChoixCod.Detail[0].PutValue('CC_LIBRE','');
    if (TobChoixCod.InsertOrUpdateDB(False) = False) then begin
      // Impossible de mettre à jour la structure de la table libre
      MsgBox.Execute(5,Caption,'');
      RollBack;
      TobCommun.Free;
      TobChoixCod.Free;
      TobNatCpte.Free;
      TobChoixExt.Free;
      BValider.Enabled := True;
      SourisNormale;
      Exit;
    end;
    AvertirTable('GCZONELIBRETIE');

    // Charge les données de la table libre source
    TobNatCpte.LoadDetailDB('NATCPTE','"T0'+cbo_CPTA.Value+'"','',nil,False,False);

    // Recopie les données dans la table destination
    for i := 0 to TobNatCpte.Detail.Count-1 do begin
      TobFille := Tob.Create('CHOIXEXT',TobChoixExt,-1);
      TobFille.PutValue('YX_TYPE',szChoixExtValue+cbo_GC.Value);
      TobFille.PutValue('YX_LIBELLE',TobNatCpte.Detail[i].GetValue('NT_LIBELLE'));
      TobFille.PutValue('YX_ABREGE',Copy(TobNatCpte.Detail[i].GetValue('NT_LIBELLE'),0,17));
      if (iLongueur <= 6) then begin	// Le code en GC est limité à 6 caractères
      	TobFille.PutValue('YX_CODE',TobNatCpte.Detail[i].GetValue('NT_NATURE'));
      	TobFille.PutValue('YX_LIBRE','');
      	end
      else begin
      	TobFille.PutValue('YX_CODE',IntToStr(i));
      	TobFille.PutValue('YX_LIBRE',TobNatCpte.Detail[i].GetValue('NT_NATURE'));
      end;
    end;
    if (TobChoixExt.InsertOrUpdateDB(False) = False) then begin
      // Impossible de mettre à jour les données de la table libre
      MsgBox.Execute(6,Caption,'');
      RollBack;
      TobCommun.Free;
      TobChoixCod.Free;
      TobNatCpte.Free;
      TobChoixExt.Free;
      BValider.Enabled := True;
      SourisNormale;
      Exit;
    end;

    // Crée les enregistrements manquants dans la table TIERSCOMPL
    ExecuteSql('INSERT INTO TIERSCOMPL (YTC_AUXILIAIRE,YTC_TIERS) SELECT T_AUXILIAIRE,T_TIERS FROM TIERS WHERE T_AUXILIAIRE NOT IN (SELECT YTC_AUXILIAIRE FROM TIERSCOMPL) AND T_TABLE'+cbo_CPTA.Value+'<>"" AND T_NATUREAUXI="'+szNature+'"');

		// Recopie des données dans la table TIERSCOMPL
    if (iLongueur <= 6) then // Le code en GC est limité à 6 caractères
      ExecuteSql('UPDATE TIERSCOMPL SET YTC_TABLELIBRE'+szChampTiersGC+cbo_GC.Value+'= (SELECT T_TABLE'+cbo_CPTA.Value+' FROM TIERS WHERE TIERSCOMPL.YTC_AUXILIAIRE = TIERS.T_AUXILIAIRE AND T_NATUREAUXI="'+szNature+'")')
    else begin
      ExecuteSql('UPDATE TIERSCOMPL SET YTC_TABLELIBRE'+szChampTiersGC+cbo_GC.Value+'= (SELECT YX_CODE FROM CHOIXEXT,TIERS WHERE YX_LIBRE = TIERS.T_TABLE'+cbo_CPTA.Value+' AND TIERS.T_AUXILIAIRE = TIERSCOMPL.YTC_AUXILIAIRE AND YX_LIBRE<>"" AND T_NATUREAUXI="'+szNature+'")');
    end;

    // Recopie des données dans la table PIECE
    // Souhaitez-vous mettre à jour les pièces en gestion commerciale ?
    if (MsgBox.Execute(4,Caption,'') = mrYes) then
      ExecuteSql('UPDATE PIECE SET GP_LIBRETIERS'+cbo_GC.Value+'= (SELECT YTC_TABLELIBRE'+szChampTiersGC+cbo_GC.Value+' FROM TIERSCOMPL WHERE PIECE.GP_TIERS = TIERSCOMPL.YTC_TIERS) WHERE PIECE.GP_TIERS IN (SELECT YTC_TIERS FROM TIERSCOMPL WHERE YTC_TABLELIBRE'+szChampTiersGC+cbo_GC.Value+'<>"")')
    else
      ExecuteSql('UPDATE PIECE SET GP_LIBRETIERS'+cbo_GC.Value+' = "" WHERE GP_LIBRETIERS'+cbo_GC.Value+' <> ""');
  end;

  TobCommun.Free;
  TobChoixCod.Free;
  TobNatCpte.Free;
  TobChoixExt.Free;
  CommitTrans;
  BValider.Enabled := True;
  SourisNormale;
  RempliCombo;
end;

{=======================================================================================}
procedure TFSynchro.FormShow(Sender: TObject);
begin
  inherited;
  HelpContext:=999999432;

  cbo_GC := cbo_GCClient;
  szChoixCodValue := 'CT';
  szChoixExtValue := 'LT';
  szChampTiersGC := 'TIERS';
  szNature := 'CLI';
  cbo_GCFournisseur.Visible := False;
  cbo_GCClient.Visible := True;
  RempliCombo;
end;

{=======================================================================================}
procedure TFSynchro.rbt_ClientClick(Sender: TObject);
begin
  inherited;
  cbo_GCClient.Visible := True;
  cbo_GCFournisseur.Visible := False;
  cbo_GC := cbo_GCClient;
  szChoixCodValue := 'CT';
  szChoixExtValue := 'LT';
  szChampTiersGC := 'TIERS';
  szNature := 'CLI';
end;

{=======================================================================================}
procedure TFSynchro.rbt_FournisseurClick(Sender: TObject);
begin
  inherited;
  cbo_GCClient.Visible := False;
  cbo_GCFournisseur.Visible := True;
  cbo_GC := cbo_GCFournisseur;
  szChoixCodValue := 'FT';
  szChoixExtValue := 'LC';
  szChampTiersGC := 'FOU';
  szNature := 'FOU';
end;

{=======================================================================================}
procedure TFSynchro.RempliCombo;
var
  TobCombo : Tob;
  i : Integer;
  szLibelle : String;
begin
  cbo_CPTA.Clear;
  TobCombo := Tob.Create('Table COMMUN',nil,-1);
  TobCombo.LoadDetailDBFromSQL('Table COMMUN','SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE = "NAT" AND CC_CODE LIKE "T0%"');
  for i := 0 to 9 do begin
    if (i<TobCombo.Detail.count) and (TobCombo.Detail.count >0)  then
      szLibelle := TobCombo.Detail[i].GetValue('CC_LIBELLE')
    else szLibelle := '';
    if (szLibelle <> '') then cbo_CPTA.Items.Add(szLibelle)
    else cbo_CPTA.Items.Add('Table n°'+IntToStr(i+1));
  end;
  FreeAndNil(TobCombo);

  cbo_GCClient.Clear;
  TobCombo := Tob.Create('Table CHOIXCOD',nil,-1);
  TobCombo.LoadDetailDBFromSQL('Table CHOIXCOD','SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE = "ZLT" AND CC_CODE LIKE "CT%"');
  for i := 0 to 9 do begin
    if (i<TobCombo.Detail.count) and (TobCombo.Detail.count >0) then
      szLibelle := TobCombo.Detail[i].GetValue('CC_LIBELLE')
    else szLibelle := '';
    if (szLibelle <> '') then cbo_GCClient.Items.Add(szLibelle)
    else cbo_GCClient.Items.Add('Table n°'+IntToStr(i+1));
  end;
  FreeAndNil(TobCombo);

  cbo_GCFournisseur.Clear;
  TobCombo := Tob.Create('Table CHOIXCOD',nil,-1);
  TobCombo.LoadDetailDBFromSQL('Table CHOIXCOD','SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE = "ZLT" AND CC_CODE LIKE "FT%"');
  for i := 0 to 2 do begin
    if (i<TobCombo.Detail.count) and (TobCombo.Detail.count >0) then
      szLibelle := TobCombo.Detail[i].GetValue('CC_LIBELLE')
    else szLibelle := '';
    if (szLibelle <> '') then cbo_GCFournisseur.Items.Add(szLibelle)
    else cbo_GCFournisseur.Items.Add('Table n°'+IntToStr(i+1));
  end;
  FreeAndNil(TobCombo);
end;

{=======================================================================================}
procedure TFSynchro.BFermeClick(Sender: TObject);
begin
  inherited;
Close ;
end;

end.
