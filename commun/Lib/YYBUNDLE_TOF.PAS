{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/05/2005
Modifié le ... : 07/04/06 par JSI
Description .. : Source TOF de la FICHE : YYSHAREBUNDLE ()
Mots clefs ... : TOF;YYSHAREBUNDLE
*****************************************************************}
Unit YYBUNDLE_TOF ;
                                                                                                               
Interface

Uses StdCtrls,
     Controls,
     Classes,
  Variants,
{$IFNDEF EAGLCLIENT}
     db,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     mul,
     Fe_main,
{$else}
     MaineAGL,
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HSysMenu,
     HTB97,
     uMultiDossierUtil,
     Graphics, Grids, Windows,
     UtilPgi,
     uMultiDossier,
     EBizUtil,
     wCommuns,
     M3FP,
     UTOF ;

Type
  tValideBundle = Class
  public
    procedure DeleteDeShare;
    procedure MajView;
    procedure EpureBase;
    procedure ValideParametrage;
    procedure EnregistreParametrage (bBaseRef : boolean);
    procedure SetMajView(bMajView : boolean);
    procedure SetCodeRegroupement(sCode : string);
    function  GetCodeRegroupement : string;
    constructor Create(CodeRegroupement:string;maTobBundle:TOB);
  private
    FCodeRegroupement : string;
    FSuppData : boolean;
    FMajView : boolean;
    TobBundle:TOB;
  end;
Type
  TOF_YYBUNDLE = Class (TOF)
   private
    TobBundle, TobGBundle : tob;
    TobTVSlave,TobTVMaster : Tob;
    GBundle : THGrid;
    TBBActuView, TBBEpurTable : TToolBarButton97;
    CbMultiDossier : THValComboBox;
    TVMaster,TVSlave : TTreeView;
    Action : TActionFiche;
    sBasesRegroupement : string;
    ValideBundle : tValideBundle;

  public
    bBasePrincipale, bModifie, bRegroupementOk : boolean;
    tnclick : TTreeNode;
    stCircularite : string;

    { affichage des bundles }
    procedure AfficheGBundle(sAffiche : string ; Arow : integer = -1);
    procedure BundleAfficheOnChange (Sender : TObject);
    procedure ChargeAndShowRegroupement(CodeMultiDossier : string);

    { gestion des 2 treeviews }
    procedure AfficheTreeViews(TobB : tob);
    procedure BVoirOnClick(Sender : TObject);
    procedure mnbOnClick(Sender : TObject);
    procedure mnhOnClick(Sender : TObject);
    procedure RedimensionnePanelsTV;
    procedure SetControlTVVisible;
    procedure SetAlignPanelTV;
    procedure SetImages (TV : TTreeView; TN : TTreeNode);
    function  CreeTobNoeuds (TobNoeud : TOB; stCode, stLibelle : string) : boolean;
    function  CreeTobNoeudm (TobNoeud : TOB; stCode, stLibelle : string) : boolean;

    { chargement des données }
    procedure chargeTobBundle(sBases : string);
    procedure ChargeTobTVSlave(TobB : tob);
    procedure ChargeTobTVMaster(TobB : tob);

    procedure ChargeLesBitmaps;
{    procedure EnregistreParametrage (bBaseRef : boolean);}

    function  GetBasesRegroupement : string;
    function  GetImage(image  : integer) : integer;
    function  GetAction : tactionfiche;

    procedure SetBasesRegroupement(LesBases : string);
    procedure SetAction(typeaction : tactionfiche);


    procedure AjouteBundle(TobB : tob);
    procedure RetireBundle(TobB : tob ; all : boolean);
    procedure partage(t : tob   ; withModifie : boolean);
    procedure departage(t : tob ; withModifie : boolean);

{    procedure DeleteDeShare;
    procedure MajView;
    procedure EpureBase;
    procedure ValideParametrage;}
    function OkConnect (S : String) : boolean;

    { évènement }
    procedure BundleOnRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure CbMultiDossierOnClick (Sender : TObject);
    procedure SplitMoved(Sender: TObject);

    { bouton }
    procedure bParamClick(Sender : TObject);
    procedure TBBActuViewOnClick (Sender : TObject);
    procedure TBBEPURTABLEOnClick (Sender : TObject);
    procedure TBBImportOnClick (Sender : TObject);
    procedure TBBExportOnClick (Sender : TObject);

    { menu pop du grid GBundle }
    procedure MnAjouteBundleClick (Sender : TObject);
    procedure MnRetireBundleClick (Sender : TObject);
    procedure MnDetailBundleClick (Sender : TObject);
    procedure PopActionOnPopup (Sender : TObject);

    { inherited }
    procedure OnLoad                   ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

  TOF_YYMAJVUEMULTISOC = Class (TOF)
  public
    procedure OnArgument (S : String) ; override;
    procedure BValiderClick(Sender: TObject);
  end;

procedure YYLanceFiche_Bundle (Nat,Cod : String ; Range,Lequel,Argument : string);
procedure MajBundlePartage(bDeleteData : boolean = false);
function CreerVueMultiSoc (NomVue, stSql : string) : boolean;
procedure UpdateMultiDossier (laBase : string; bConserveGuid : boolean = true) ;

Implementation

uses Math, hpanel, majtable, BundleDetail_TOF,Menus,
     SaisUtil, ParamSoc;

Const
  { images }
  imPartage = 44;
  imNonDefini = 41;
  imNonPartage = 43;
  imModifie = 49;

  { code du bundle Regroupements dossiers' }
  sCodeRDoss = 'YS3';


function tValideBundle.GetCodeRegroupement :string;
begin
  result := FCodeRegroupement;
end;

procedure tValideBundle.SetCodeRegroupement(sCode : string);
begin
  FCodeRegroupement := sCode;
end;

procedure tValideBundle.SetMajView(bMajView : boolean);
begin
  FMajView := bMajView;
end;

constructor tValideBundle.Create(CodeRegroupement:string;maTobBundle:TOB);
begin
  if (CodeRegroupement <> '')
    then FCodeRegroupement:=CodeRegroupement
    else FCodeRegroupement:=MS_CODEREGROUPEMENT;
  FSuppData:=False;
  FMajView := True;
  TobBundle:=maTobBundle;
end;

Procedure YYLanceFiche_Bundle (Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
  AGLLanceFiche (Nat,Cod,Range,Lequel,Argument);
end;

{ Affichage }
procedure TOF_YYBUNDLE.AfficheGBundle(sAffiche : string ; Arow : integer = -1);
var TobG, TobRow : tob;
    iInd : integer;
begin
  TobGBundle.ClearDetail;
  if assigned(GBundle) then
  begin
    if sAffiche = '' then
    begin
      TobGBundle.Dupliquer(TobBundle,true,true);
    end else
    begin
      for iInd := 0 to TobBundle.Detail.Count -1 do
      begin
        { BP = bundle partagé }  { NBP = non partagé }
        if (TobBundle.Detail[iInd].GetBoolean('PARTAGE') and (sAffiche = 'BP'))
        or (not TobBundle.Detail[iInd].GetBoolean('PARTAGE') and (sAffiche = 'BNP')) then
        begin
          TobG := tob.create('',TobGBundle,-1);
          TobG.Dupliquer(TobBundle.Detail[iInd],true,true);
        end;
      end;
    end;

   TobGBundle.PutGridDetail (GBundle, False, False, 'CO_ABREGE;CO_LIBELLE;IM_MODIFIE;IM_PARTAGE', True);
    { code }
    GBundle.ColWidths[1] := 70;
    { libellé }
    GBundle.ColWidths[2] := 120;
    { flag modifié }
    GBundle.ColWidths[3] := 20;
    { flag partagé }
    GBundle.ColWidths[4] := 30;

    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns (GBundle);

    { on se positionne sur la ligne passée en paramètre }
    if Arow <> -1 then
    begin
      if ARow <= TobGBundle.Detail.Count-1 then
      begin
        GBundle.Row := Arow;
        TobRow := Tob(GBundle.Objects [0, Arow]);
        AfficheTreeViews(TobRow);
      end;
    end else
    begin
      if TobGBundle.Detail.Count > 0 then
        AfficheTreeViews(TobGBundle.Detail[0])
      else AfficheTreeViews(nil);
    end;
  end;
end;

function TOF_YYBUNDLE.CreeTobNoeuds (TobNoeud : TOB; stCode, stLibelle : string) : boolean;
var
  TobB, TobN : tob;
  iInd : integer;
begin
  result := (stcode <> '') or (not assigned (TobNoeud.FindFirst(['CODE'],[stCode],true)));
  if result then
  begin
    if stcode <> '' then
    begin
      TobN := Tob.Create ('', TobNoeud, -1);
      TobN.AddChampSupValeur ('CODE', stCode);
      TobN.AddChampSupValeur ('LIBELLE', stLibelle);
    end else TobN := TobNoeud;

    TobB := TobBundle.FindFirst (['CO_CODE'], [TobN.GetString ('CODE')], false);
    if assigned(TobB) then
    begin
      { lecture des bundles associés }
      for iInd := 0 to TobB.Detail.Count - 1 do
      begin
        Result := CreeTobNoeuds (TobN, TobB.Detail[iInd].GetString ('CO_CODE'), TobB.Detail[iInd].GetString ('CO_LIBELLE'));
        if not result then break;
      end;
    end;
  end else stCircularite := stcode;
end;

function TOF_YYBUNDLE.CreeTobNoeudm (TobNoeud : TOB; stCode, stLibelle : string) : boolean;
var
  TobN : tob;
  iInd : integer;
begin
  result := (stcode <> '') or (not assigned (TobNoeud.FindFirst(['CODE'],[stCode],true)));
  if result then
  begin
    if stcode <> '' then
    begin
      TobN := Tob.Create ('', TobNoeud, -1);
      TobN.AddChampSupValeur ('CODE', stCode);
      TobN.AddChampSupValeur ('LIBELLE', stLibelle);
    end else TobN := TobNoeud;
    { lecture des bundles }
    for iInd := 0 to TobBundle.Detail.Count - 1 do
    begin
      if assigned(TobBundle.Detail[iInd].FindFirst (['CO_CODE'], [TobN.GetString ('CODE')], false)) then
      begin
        Result := CreeTobNoeudm(TobN, TobBundle.Detail[iInd].GetString ('CO_CODE'), TobBundle.Detail[iInd].GetString ('CO_LIBELLE'));
        if not result then break;
      end;
    end;
  end else stCircularite := stcode;
end;

procedure TOF_YYBUNDLE.SetImages (TV : TTreeView; TN : TTreeNode);
var
  tn1 : TTreeNode;
  numImage : integer;
  TobS,TobData : TOB;
  sCode : string;
begin
  if not assigned(TV) and  not assigned(TN) then Exit;
  if assigned(TV) then
  begin
    if TV.Items.count > 0 then tn1 := TV.Items [0] //TopItem
    else tn1 := nil;
  end else if TN.Count <> 0 then tn1 := TN.getFirstChild
  else tn1 := TN;

  while assigned(tn1) do
  begin
    NumImage := 0;
    TobData := Tob(tn1.Data);
    if assigned (TobData) then
    begin
      sCode := Tob(tn1.Data).GetString ('CODE');
      if sCode <> '' then
      begin
        TobS := TobBundle.FindFirst (['CO_CODE'], [sCode], false);
        if assigned(TobS) and TobS.GetBoolean('PARTAGE') then NumImage := GetImage(ImPartage)
        else NumImage := GetImage(ImNonPartage);
      end;  
    end;
    tn1.ImageIndex := NumImage -1;
    tn1.SelectedIndex := tn1.ImageIndex;
    if tn1.Count > 0 then SetImages(nil, tn1);
    if (assigned(TV)) or (TN.Count <> 0) then tn1 := tn1.GetNextSibling else tn1 := nil;
  end;

  TVSlave.Invalidate;
end;

procedure TOF_YYBUNDLE.AfficheTreeViews(TobB : tob);
var
  iNiv : integer;
  stCol : string;
begin

  { on charge les données }
  ChargeTobTVSlave(TobB);
  ChargeTobTVMaster(TobB);

  { maj de l'interface}
  //if TobTVSlave.Detail.Count > 0 then
  begin
    stCol := '';
    for iNiv := 0 to TobTVSlave.MaxNiveau do
      stCol := stCol + 'LIBELLE;';

    TobTVSlave.PutTreeView (TVSlave, nil, stCol);
    SetImages (TVSlave, nil);
    TVSlave.FullExpand;
  end;

  { maj de l'interface}
  //if TobTVMaster.Detail.Count > 0 then
  begin
    stCol := '';
    for iNiv := 0 to TobTVMaster.MaxNiveau do
      stCol := stCol + 'LIBELLE;';

    TobTVMaster.PutTreeView (TVMaster, nil, stCol);
    SetImages (TVMaster, nil);
    TVMaster.FullExpand;
  end;
end;

procedure TOF_YYBUNDLE.BVoirOnClick(Sender : TObject);
begin
  SetControlChecked ('MNH', GetControlVisible('PH'));
  SetControlChecked ('MNB', GetControlVisible('PB'));
end;

procedure TOF_YYBUNDLE.mnbOnClick(Sender : TObject);
begin
  SetControlChecked ('MNB', not (GetCheckBoxState ('MNB')=cbchecked));
  SetControlVisible('PB', not GetControlVisible('PB'));
  SetAlignPanelTV;
  RedimensionnePanelsTV;
  SetControlTVVisible;
  if GetControlVisible('SPLITVERT') then
    ThSplitter(GetControl('SPLITVERT')).Left := ThPanel(GetControl('PRIGHT')).Left - ThSplitter(GetControl('SPLITVERT')).Width -1;
end;

procedure TOF_YYBUNDLE.mnhOnClick(Sender : TObject);
begin
  SetControlChecked ('MNH', not (GetCheckBoxState ('MNH')=cbchecked));
  SetControlVisible('PH', not GetControlVisible('PH'));
  SetAlignPanelTV;
  RedimensionnePanelsTV;
  SetControlTVVisible;
  if GetControlVisible('SPLITVERT') then
    ThSplitter(GetControl('SPLITVERT')).Left := ThPanel(GetControl('PRIGHT')).Left - ThSplitter(GetControl('SPLITVERT')).Width -1;
end;

procedure TOF_YYBUNDLE.RedimensionnePanelsTV;
var PH, PB, PRight : ThPanel;
begin
  PH := ThPanel(GetControl('PH'));
  PB := ThPanel(GetControl('PB'));
  PRight := ThPanel(GetControl('PRIGHT'));

  if PB.Visible and PH.Visible then
  begin
    PH.Height :=  trunc(PRight.Height / 2);
    PB.Height :=  PRight.Height-PH.Height;
    ThSplitter(GetControl('SPLITHORTZ')).Top := PB.Top + ThSplitter(GetControl('SPLITHORTZ')).Height +1;
  end;
end;

procedure TOF_YYBUNDLE.SetControlTVVisible;
var bVisible : boolean;
begin
  bVisible := (GetControlVisible('PB') or GetControlVisible('PH'));
  if bVisible <> GetControlVisible('PRIGHT') then
  begin
    SetControlVisible('PRIGHT', bVisible);
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns (GBundle);
  end;
  SetControlVisible('SPLITVERT', GetControlVisible('PRIGHT'));
  SetControlVisible('SPLITHORTZ', (GetControlVisible('PB') and GetControlVisible('PH')));
end;

procedure TOF_YYBUNDLE.SetAlignPanelTV;
var PH, PB : ThPanel;
begin
  PH := ThPanel(GetControl('PH'));
  PB := ThPanel(GetControl('PB'));

  if PH.Visible and not PB.Visible then
  begin
    PB.Align := alTop;
    PH.Align := alClient;
  end else
  if PB.Visible and not PH.Visible then
  begin
    PH.Align := alTop;
    PB.Align := alClient;
  end else
  if PB.Visible and PH.Visible then
  begin
    PH.Align := alTop;
    PB.Align := alClient;
  end;
end;

procedure UpdateMultiDossier (laBase : string; bConserveGuid : boolean = true) ;
var t : tob;
    Qdoss : tQuery;

  function getNoDossier() : string;
  var iDoss :  integer;
      Qmax  : tQuery;
  begin
    iDoss := 0;
    Qmax := OpenSQL ('SELECT MAX(DOS_NODOSSIER) FROM DOSSIER', true);
    try
      if (not Qmax.Eof) and (not VarIsNull(Qmax.Fields[0].Value)) then
      begin
        if isNumeric(Qmax.Fields[0].AsString) then
          iDoss := StrToInt(Qmax.Fields[0].AsString)
        else iDoss := -1;
      end;
    finally
      ferme(Qmax);
    end;
    if iDoss >= 0 then
      result := Format('%.8d', [iDoss + 1])
    else result := '';
  end;

  function getGuidDossier() : string;
  begin
    result := AglGetGuid();
  end;

  function GetMonParamsoc (LaBase, LeNom : string) : string;
  var
    Q : TQuery;
  begin
    Q := OpenSql ('SELECT SOC_DATA FROM ' + GetBase (LaBase, 'PARAMSOC')
              + ' WHERE SOC_NOM="' + LeNom + '"', true);
    if not Q.Eof then
      result := Q.Fields[0].AsString
    else Result := '';
    Ferme (Q);
  end;

  procedure updateParamSoc(noDoss, guidDoss : string);
  begin
    executeSql('UPDATE ' + GetBase (LaBase, 'PARAMSOC')
              + ' SET SOC_DATA="' + noDoss + '"'
              + ' WHERE SOC_NOM="SO_NODOSSIER"');

    executeSql('UPDATE ' + GetBase (LaBase, 'PARAMSOC')
              + ' SET SOC_DATA="' + guidDoss + '"'
              + ' WHERE SOC_NOM="SO_GUIDDOSSIER"');
  end;

begin
  Qdoss := openSql('SELECT * FROM DOSSIER' +
                  ' WHERE DOS_NOMBASE="' + laBase + '"', true);
  try
    { création d'un dossier }
    if Qdoss.Eof then
    begin
      t := tob.Create('DOSSIER',nil,-1);
      try
        t.InitValeurs();
        t.SetString('DOS_NODOSSIER',getNoDossier);
        t.SetString('DOS_GUIDDOSSIER',getGuidDossier);
        t.SetString('DOS_NOMBASE',laBase);
        t.SetString('DOS_LIBELLE',laBase); //js1 18102006 pour Jean Pasteris        
        if t.GetString('DOS_NODOSSIER') <> '' then
        begin
          if t.InsertDb(nil) then
          begin
            updateParamsoc(t.GetString('DOS_NODOSSIER'),t.GetString('DOS_GUIDDOSSIER'));
          end;
        end;  
      finally;
        t.Free;
      end;
    end else
   { on vérifie que le GUID est renseigné }
    begin
      t := tob.Create('DOSSIER',nil,-1);
      try
        t.SelectDB('',Qdoss);
        if (t.GetString('DOS_GUIDDOSSIER') = '') or    { PCL }
           (t.GetString('DOS_GUIDDOSSIER') <> getMonParamSoc (LaBase, 'SO_GUIDDOSSIER')) or
           (t.GetString('DOS_NODOSSIER')   <> getMonParamSoc (LaBase, 'SO_NODOSSIER'))   then
        begin
          if not bConserveGuid then t.SetString('DOS_GUIDDOSSIER',getGuidDossier);
          if t.UpdateDB() then
          begin
            updateParamsoc(t.GetString('DOS_NODOSSIER'),t.GetString('DOS_GUIDDOSSIER'));
          end;
        end;
      finally
        t.Free;
      end;
    end;
  finally
    ferme(Qdoss);
  end;
end;

procedure tValideBundle.EnregistreParametrage (bBaseRef : boolean);
var
  UneBase, stLesBases ,sttypebase : string;
  TobShare : tob;

  procedure UpdateDeShare;
  var
    iInd, iIndEB : integer;
    TobBD : Tob;
    stSql : string;
  begin
    for iInd := 0 to TobShare.Detail.Count - 1 do
    begin
      for iIndEB := 0 to TobShare.Detail[iInd].Detail.Count - 1 do
      begin
        TobBD := TobShare.Detail[iInd].Detail[iIndEB];
        if pos (TobBD.GetString ('CO_ABREGE'), 'TAB;TTE;PAR') > 0 then
        begin
          // Test l'existence de l'enregistrement avant INSERT sinon violation de clé
          if not ExisteSQL('SELECT DS_NOMTABLE FROM ' + GetBase (UneBase, 'DESHARE') +
                             ' WHERE DS_NOMTABLE = "' + TobBD.GetString ('CO_LIBELLE') + '"' +
                             ' AND DS_MODEFONC = "LIB"') then
          begin
            // MAJ base
            if ChampToNum ('DS_TYPTABLE') >= 0 then
            begin
              if ChampToNum ('DS_VUE') >= 0 then
              begin
                stSql := 'INSERT INTO ' + GetBase (UneBase, 'DESHARE') +
                            ' (DS_NOMTABLE, DS_MODEFONC, DS_TYPTABLE, DS_NOMBASE, DS_VUE) VALUES ("' +
                            TobBD.GetString ('CO_LIBELLE') + '", "LIB", "' + TobBD.GetString ('CO_ABREGE') + '", "' +
                            V_PGI.SchemaName + '", "' + TobBD.GetString ('VUE') + '")'
              end else
              begin
                stSql := 'INSERT INTO ' + GetBase (UneBase, 'DESHARE') + ' (DS_NOMTABLE, DS_MODEFONC, DS_TYPTABLE, DS_NOMBASE) VALUES ("' +
                            TobBD.GetString ('CO_LIBELLE') + '", "LIB", "' + TobBD.GetString ('CO_ABREGE') + '", "' + V_PGI.SchemaName + '")'
              end;
            end else
            begin
              if ChampToNum ('DS_VUE') >= 0 then
              begin
                stSql := 'INSERT INTO ' + GetBase (UneBase, 'DESHARE') + ' (DS_NOMTABLE, DS_MODEFONC, DS_NOMBASE, DS_VUE) VALUES ("' +
                            TobBD.GetString ('CO_LIBELLE') + '", "LIB", "' + V_PGI.SchemaName + '", "' + TobBD.GetString ('VUE') + '")';
              end else
              begin
                stSql := 'INSERT INTO ' + GetBase (UneBase, 'DESHARE') + ' (DS_NOMTABLE, DS_MODEFONC, DS_NOMBASE) VALUES ("' +
                            TobBD.GetString ('CO_LIBELLE') + '", "LIB", "' + V_PGI.SchemaName + '")';
              end;
            end;
            ExecuteSql (stSql);
          end;
        end;
      end;
    end;
  end;

  procedure chargeLesDonnees();
  var iInd, jInd : integer;
      TobDetail, t : tob;
      sLib, sTable, sVue : string;
  begin
     for iInd := 0 to TobBundle.Detail.Count -1 do
     begin
       if TobBundle.Detail[iInd].GetBoolean('PARTAGE') then
       begin
         TobDetail := tob.Create('',TobShare,-1);
         TobDetail.LoadDetailDBFromSql ('', 'SELECT * FROM COMMUN WHERE CO_TYPE="YEB" AND CO_LIBRE="' + TobBundle.Detail[iInd].GetString ('CO_CODE') + '"');

         { gestion des vues : le CO_LIBELLE peut contenir TABLE;VUE }
         for jInd := 0 to TobDetail.Detail.Count -1 do
         begin
           t := TobDetail.Detail[jInd];
           t.AddChampSupValeur('VUE','');
           if t.GetString ('CO_ABREGE') = 'TAB' then
           begin
             sLib := t.GetString ('CO_LIBELLE');
             sTable := ReadTokenSt (sLib);
             sVue := ReadTokenSt (sLib);
             if sVue <> '' then t.SetString ('VUE', sVue);
             t.SetString ('CO_LIBELLE', sTable);
           end;
         end;

       end;
     end;
  end;

begin
  if FCodeRegroupement <> '' then
  begin
    { chargement de la tob deshare }
    TobShare := tob.Create('',nil,-1);
    try
      chargeLesDonnees();
      // Récupération des bases de regroupement
      stLesBases := GetBasesMS (FCodeRegroupement, true);

      // Remplissage table temporaire pour chaque base :
      UneBase := ReadTokenSt (stLesBases) ;
      while UneBase <> '' do
      begin
        if bBaseRef then sttypebase:='la base commune' else sttypebase:='la(les) base(s) dépendante(s)';
        wMoveProgressForm ('Enregistrement dans '+sttypebase+' : ' + UneBase);
        if ((UneBase = V_PGI.SchemaName) and bBaseRef) or
           ((UneBase <> V_PGI.SchemaName) and not bBaseRef) then
          UpdateDeShare;
          UpdateMultiDossier (UneBase, false);
        UneBase := ReadTokenSt (stLesBases) ;
      end ;
    finally
      TobShare.Free;
    end;
  end
  else
    PgiBox ('Le Regroupement multisociété n''a pas été choisi');

//  bModifie := false;
end;

procedure tValideBundle.DeleteDeShare;
var
  stLesBases, UneBase : string;
begin
  wMoveProgressForm ('Annulation ancien paramétrage');

  stLesBases := GetBasesMS (FCodeRegroupement, true);
  UneBase := ReadTokenSt (stLesBases) ;

  while UneBase <> '' do
  begin
    ExecuteSql ('DELETE FROM ' + GetBase (UneBase, 'DESHARE'));
    UneBase := ReadTokenSt (stLesBases) ;
  end;
end;

{ évènement }

procedure TOF_YYBUNDLE.BundleOnRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  AfficheTreeViews(Tob(GBundle.Objects [0, Ou]));
end;

procedure TOF_YYBUNDLE.TBBActuViewOnClick (Sender: TObject);
begin
  Transactions (ValideBundle.MajView, 1);
end;

procedure tValideBundle.MajView;
var
  StDriver, StServer, StPath, Stbase, StUser, StPassWord, StODBC, StOptions, StGroup, StShare: string;
  DB : TDataBase;
  stLesBases, laSoc, SchemaName : string;
  bSaveEnableDeShare, bSaveExisteDeShare : boolean;
begin
  SchemaName := V_PGI.SchemaName;

  DBDeleteAllView(DBSOC, Nil, V_PGI.Driver);
  DBCreateAllView(DBSOC, Nil, V_PGI.Driver);

  bSaveEnableDeShare := V_PGI.enableDEShare;
  bSaveExisteDeShare := V_PGI.existeDEShare;

  if  ExisteSql ('SELECT 1 FROM DESHARE WHERE DS_NOMBASE="' + V_PGI.SchemaName + '"') then
  begin
    stLesBases := GetBasesMS (FCodeRegroupement, false);
    laSoc := ReadTokenSt (stLesBases) ;

    DB := TDataBase.Create(Appli);
    while laSoc <> '' do
    begin
      if laSoc <> V_PGI.FCurrentAlias then
      begin
        wMoveProgressForm ('Mise à jour des vues de la base ' + LaSoc);
        PutToDebuglog ('Mise à jour des vues de la base ' + LaSoc +' à '+TimeToStr(Now),True);        
        ChargeDBParams (laSoc, StDriver, StServer, StPath, Stbase, StUser, StPassWord,
                        StODBC, StOptions, StGroup, StShare);
        AssignDBParamsStrings ('DB001', StDriver, StServer, StPath, StBase, StUser, StPassWord,
                               StODBC, StOptions, DB);
        DB.Connected := TRUE;
        ChargeDeShare ();

        V_PGI.enableDEShare := true;
        V_PGI.existeDEShare := true;

        MajOracleSession(DB);

        DBDeleteAllView(DB, Nil, V_PGI.Driver);
        DBCreateAllView(DB, Nil, V_PGI.Driver);

        DB.connected := false;
      end;
      laSoc := ReadTokenSt (stLesBases);
    end;
    freeandNil(DB);
  end;
  V_PGI.enableDEShare := bSaveEnableDeShare;
  V_PGI.existeDEShare := bSaveExisteDeShare;
end;


procedure TOF_YYBUNDLE.CbMultiDossierOnClick (Sender : TObject);
var
  CodeMultiDossier, stLesBases  : string;
begin
  bRegroupementOk := (trim(ThValComboBox(Sender).Text) <> '');
  CodeMultiDossier := ThValComboBox(Sender).Value;
  stLesBases := GetBasesMS (CodeMultiDossier, true);
  if Pos (V_PGI.SchemaName, stLesBases) = 0 then
  begin
    bRegroupementOk := false;
    PGIBox ('La base connectée ne fait pas partie du regroupement');
    ChargeAndShowRegroupement('');
  end
  else
  begin
    ChargeAndShowRegroupement(CodeMultiDossier);
  end;
end;

procedure TOF_YYBUNDLE.TBBEPURTABLEOnClick (Sender : TObject);
begin
  Transactions (ValideBundle.EpureBase, 1);
end;

procedure TOF_YYBUNDLE.bParamClick (Sender : TObject);
begin
  ParamRegroupementMultidossier;
end;

procedure TOF_YYBUNDLE.TBBExportOnClick (Sender : TObject);
var
  stFile : string;
  TobShare, tobS, t : tob;
  iInd : integer;
begin
  stFile := ExtractFileDir(Application.ExeName);
  TobShare := tob.Create('', nil,- 1);
  try
    for iInd := 0 to TobBundle.Detail.Count -1 do
    begin
      t := TobBundle.Detail[iInd];
      if t.GetBoolean('PARTAGE') then
      begin
        tobS := tob.create('COMMUN',TobShare,-1);
        tobS.Dupliquer(t,true,true);
      end;
    end;
    TOBToXXXFile(TobShare, stFile + '\MULTISOC.bin', true);
    TobShare.SaveToXmlFile (stFile + '\MULTISOC.xml', true, true);
    PgiBox ('La sauvegarde a été effectuée', Ecran.Caption);
  finally
    TobShare.Free;
  end;
end;

procedure TOF_YYBUNDLE.TBBImportOnClick (Sender : TObject);
var
  stFile : string;
  TobF, TobFile, TobB : TOB;
  iInd, iInd2 : integer;
begin
  stFile := ExtractFileDir(Application.ExeName);
  TobFile := Tob.Create ('', nil, -1);
  try
    XXXFileToTOB (stFile + '\MULTISOC.bin', TobFile, true);
    //TobBundle.ClearDetail;
    if TobBundle.Detail.Count > 0 then
    begin
      TobBundle.PutValueAllFille('PARTAGE',false);
      TobBundle.PutValueAllFille('PARTAGESTR','-');
      TobBundle.PutValueAllFille('IM_PARTAGE','#ICO#'+intToStr(GetImage(imNonPartage)));
      TobBundle.PutValueAllFille('MODIFIE',false);
      TobBundle.PutValueAllFille('IM_MODIFIE','');
    end;
    for iInd := TobFile.Detail.Count - 1 downto 0 do
    begin
      for iInd2 := TobFile.Detail[iInd].Detail.Count - 1 downto 0 do
      begin
        TobF := TobFile.Detail[iInd].Detail[iInd2];
        TobB := TobBundle.FindFirst(['CO_CODE'],[TobF.GetString('CO_CODE')],false);
        if assigned(TobB) then
        begin
          partage(TobB,false);
        end;
      end;
    end;
  finally
    TobFile.Free;
  end;

  AfficheGBundle(GetControlText('BUNDLEAFFICHE'));

  PgiBox ('La restauration a été effectuée', Ecran.Caption);
end;

procedure TOF_YYBUNDLE.SplitMoved(Sender: TObject);
begin
  THSystemMenu(GetControl('HMTrad')).ResizeGridColumns (GBundle);
end;

procedure TOF_YYBUNDLE.BundleAfficheOnChange (Sender : TObject);
begin
  AfficheGBundle(GetControlText('BUNDLEAFFICHE'));
end;

procedure TOF_YYBUNDLE.MnAjouteBundleClick (Sender : TObject);
var TobB : tob;
    Arow : integer;
    all : boolean;

  function ExisteParentNonPartage(t : tob) : boolean;
  var
      iInd : integer;
  begin
    result := false;
    for iInd := 0 to t.Detail.Count -1 do
    begin
       if not result then
         result := assigned(TobBundle.FindFirst(['CO_CODE','PARTAGESTR'],[t.Detail[iInd].GetString('CO_CODE'),'-'],false))
       else break;
    end;
  end;

begin
  Arow := GBundle.Row;
  all := true;
  TobB := Tob(GBundle.Objects [0, Arow]);
  if assigned(TobB) then
  begin
    if ExisteParentNonPartage(TobB) then
      case PgiAskCancel('Cette action entrainera le partage d''autres bundles,#13 voulez vous continuer ?') of
        mrYes : all := true;
        mrNo :  all := false;
        else all := false;
      end;

    if all then
    begin
      AjouteBundle(TobBundle.FindFirst(['CO_CODE'],[TobB.GetString('CO_CODE')],false));
      AfficheGBundle(GetControlText('BUNDLEAFFICHE'),Arow);
    end;
  end;  
end;

procedure TOF_YYBUNDLE.MnRetireBundleClick (Sender : TObject);
var
    TobB : tob;
    Arow : integer;
    all : boolean;

    function ExisteBundleDependant() : boolean;
    var
        iInd : integer;
    begin
      result := false;
      for iInd := 0 to TobTVMaster.Detail.Count -1 do
      begin
         if not result then
           result := assigned(TobBundle.FindFirst(['CO_CODE','PARTAGESTR'],[TobTVMaster.Detail[iInd].GetString('CODE'),'X'],false))
         else break;
      end;
    end;

begin
  Arow := GBundle.Row;
  TobB := Tob(GBundle.Objects [0, Arow]);
  if assigned(TobB) then
    TobB := TobBundle.FindFirst(['CO_CODE'],[TobB.GetString('CO_CODE')],false);
  if assigned(TobB) then
  begin
    if TobB.GetString('CO_CODE') = sCodeRDoss then
    begin
      PgiBox('Le bundle ''Regroupements dossiers'' n''est pas modifiable',Ecran.Caption);
    end else
    begin
      if TobB.GetBoolean('PARTAGE') then
      begin
        all := true;
        if ExisteBundleDependant() then
        begin
          case PgiAskCancel('Cette action entrainera le retrait du partage d''autres bundles,#13 voulez vous continuer ?') of
            mrYes : all := true;
            mrNo :  all := false;
            else all := false;
          end;
        end;
        if all then
        begin
          RetireBundle(TobB, true);
          AfficheGBundle(GetControlText('BUNDLEAFFICHE'), Arow);
        end;
      end;
    end;
  end;
end;

procedure TOF_YYBUNDLE.MnDetailBundleClick (Sender : TObject);
begin
  if assigned(Tob(GBundle.Objects [0, GBundle.Row])) then
    GCLanceFiche_BundleDetail('YY', 'YYBUNDLEDETAIL', '', '', Tob(GBundle.Objects [0, GBundle.Row]).GetString('CO_CODE'));
end;

procedure TOF_YYBUNDLE.PopActionOnPopup (Sender : TObject);
var
    TobB : tob;
    isPartage : boolean;
begin
  TobB := Tob(GBundle.Objects [0, GBundle.Row]);
  if assigned(TobB) then
  begin
    isPartage := TobB.GetBoolean('PARTAGE');
    SetControlEnabled('MNAJOUTE',not isPartage);
    SetControlEnabled('MNRETIRE',isPartage);
  end;
end;

procedure TOF_YYBUNDLE.AjouteBundle (TobB : tob);
var
  iInd : integer;
  TobF : Tob;

  function EstPartage(t : tob) : boolean;
  var TobP : tob;
  begin
    result := false;
    if assigned (t) then
    begin
      TobP := TobBundle.FindFirst(['CO_CODE'],[t.GetString('CO_CODE')],false);
      if assigned (TobP) then result := TobP.GetBoolean('PARTAGE');
    end;
  end;

begin
  if assigned(TobB) then
  begin
    if not TobB.GetBoolean('PARTAGE') then
    //begin
      partage(TobB,true);
    for iInd := 0 to TobB.Detail.Count -1 do
    begin
      //if not EstPartage(TobB.Detail[iInd]) then
      //begin
        TobF := TobBundle.FindFirst(['CO_CODE'],[TobB.Detail[iInd].GetString('CO_CODE')],false);
        if assigned(TobF) then AjouteBundle(TobF)
        else if not EstPartage(TobB.Detail[iInd]) then partage(TobB.Detail[iInd],true);
      //end;
    end;
    //end;
  end;
end;

procedure TOF_YYBUNDLE.RetireBundle(TobB : tob ; all : boolean);
var
  iInd : integer;

  procedure SetDePartageChampSup(t : tob);
  begin
    if t.GetString('CO_CODE') <> sCodeRDoss then
    begin
      departage(t,true);
    end;
  end;

  function EstPartage(t : tob) : boolean;
  var TobP : tob;
  begin
    result := false;
    if assigned (t) then
    begin
      TobP := TobBundle.FindFirst(['CO_CODE'],[t.GetString('CO_CODE')],false);
      if assigned (TobP) then result := TobP.GetBoolean('PARTAGE');
    end;
  end;

begin
  if assigned(TobB) then
  begin
    if TobB.GetBoolean('PARTAGE') then
    //begin
      SetDePartageChampSup(TobB);
      for iInd := 0 to TobBundle.Detail.Count -1 do
      begin
        if TobBundle.Detail[iInd].GetBoolean('PARTAGE') then
        begin
          if assigned(TobBundle.Detail[iInd].FindFirst(['CO_CODE'],[TobB.GetString('CO_CODE')],false)) then
            RetireBundle(TobBundle.Detail[iInd],all);
        end;//SetDePartageChampSup(TobBundle.Detail[iInd]);
      end;
    //end;
  end;
end;

procedure tValideBundle.EpureBase;
var
  TobS : Tob;
  bOk : boolean;
  UneBase, stLesBases : string;

  procedure EpureTable;
  var
    iInd : integer;
  begin
    while UneBase <> '' do
    begin
      if UneBase <> V_PGI.SchemaName then
      begin
        for iInd := 0 to TobS.Detail.Count - 1 do
        begin
          // le @@ pour ne pas traduire la requete et donc ne pas utiliser DESHARE
          ExecuteSql ('@@DELETE FROM ' + GetBase (UneBase, TobS.Detail[iInd].GetString ('DS_NOMTABLE')));
        end;
        ExecuteSql ('@@DELETE FROM ' + GetBase (UneBase, 'CHOIXCOD') + ' WHERE CC_TYPE IN (SELECT DO_TYPE FROM DECOMBOS WHERE ' +
                      'EXISTS (SELECT 1 FROM DESHARE WHERE DS_TYPTABLE="TTE" AND DS_NOMTABLE=DO_COMBO))');
      end;
      UneBase := ReadTokenSt (stLesBases) ;
    end;
  end;

begin
  if FCodeRegroupement <> '' then
  begin
    if FSuppData then
    begin
      if PgiAsk ('Le paramétrage a été modifié'#13'Confirmer la suppression des données') = mrYes then
        bOk := true
      else bOk := false;
    end else bOk := true;

    if bOk then
    begin
      stLesBases := GetBasesMS (FCodeRegroupement, true);

      UneBase := ReadTokenSt (stLesBases) ;
      if UneBase <> '' then
      begin
        TobS := Tob.Create ('', nil, -1);
        TobS.LoadDetailDBFromSQL ('DESHARE', 'SELECT * FROM DESHARE WHERE DS_TYPTABLE="TAB"');
        V_PGI.IoError := oeOk;

        EpureTable;

        TobS.Free;
      end;
    end;
  end;
end;

procedure TOF_YYBUNDLE.ChargeAndShowRegroupement(CodeMultiDossier : string);
var
    ExistePartage : boolean;
begin
  if ((CodeMultiDossier = '') or not bBasePrincipale or not bRegroupementOk) then
    SetAction(taConsult)
  else SetAction(tamodif);

  if CodeMultiDossier = '' then  SetBasesRegroupement('')
  else SetBasesRegroupement(GetBasesMS(CodeMultiDossier, true));
  chargeTobBundle(GetBasesRegroupement);
  AfficheGBundle(GetControlText('BUNDLEAFFICHE'));

  ExistePartage := assigned(TobBundle.FindFirst (['PARTAGESTR'], ['X'], false));

  { on active/désactive les options }
  SetControlVisible('MNAJOUTE', (GetAction = tamodif));
  SetControlVisible('MNRETIRE', (GetAction = tamodif));
  SetControlVisible('TBBACTUVIEW', (not bBasePrincipale) and bRegroupementOk and ExistePartage);
  SetControlVisible('TBBEPURTABLE', (not bBasePrincipale) and (V_PGI.SAV) and bRegroupementOk and ExistePartage);
  SetControlEnabled('TBBEXPORT', bBasePrincipale and bRegroupementOk and ExistePartage);
  SetControlEnabled('TBBIMPORT', bBasePrincipale and bRegroupementOk);
  SetControlEnabled('bValider', bBasePrincipale and bRegroupementOk);
end;


{ inherited }
procedure TOF_YYBUNDLE.OnLoad ;
begin
  Inherited ;
  if ExisteSql ('SELECT 1 FROM DESHARE WHERE DS_NOMBASE<>"' + V_PGI.SchemaName + '"') then
  begin
    PgiBox ('La base partage des éléments avec une autre base de référence'#13'Le paramétrage ne pourra être enregistré');
    bBasePrincipale := false;
  end else
  begin
    {$IFDEF PGIMAJVER}
    bBasePrincipale := true;
    CbMultiDossier.Plus := ' YMD_CODE<>"' + MS_CODEREGROUPEMENT + '" ';
    {$ELSE PGIMAJVER}
    bBasePrincipale := false;
    {$ENDIF PGIMAJVER}
  end;

  if not bBasePrincipale then
  begin
    CbMultiDossier.Value := MS_CODEREGROUPEMENT;
    CbMultiDossier.Enabled := false;
  end;

  ChargeAndShowRegroupement(CbMultiDossier.Value);
  SetControlText('BUNDLEAFFICHE','');
end ;

procedure TOF_YYBUNDLE.OnUpdate ;
var stLesBases : String;
begin
  if PgiAsk ('Confirmez vous l''enregistrement du paramétrage ?') = mrYes then
  begin
    if GetControlText('MULTIDOSSIER') <> '' then
    begin
      // Récupération des bases de regroupement
      stLesBases := GetBasesMS (CbMultiDossier.Value, true);

      // js1 28122006 on teste la connexion (si pas possible ==> on sort
      // pour éviter la catastrophe)
      if not OkConnect(stLesBases) then abort;
    end;

    ValideBundle.SetCodeRegroupement(GetControlText('MULTIDOSSIER'));
    ValideBundle.SetMajView(TCheckBox(GetControl('CBMAJVIEW')).Checked);
    Transactions (ValideBundle.ValideParametrage, 1);

    if V_PGI.IoError <> oeOk then
      PgiBox ('Erreur lors de l''enregistrement du paramétrage')
    else PgiBox ('Le paramétrage a bien été enregistré'#13'Veuillez vous connecter à nouveau pour prendre en compte le paramétrage');
  end;
end;

procedure TOF_YYBUNDLE.chargeTobBundle(sBases : string);
var
  iInd : integer;
  TobB  : tob;

  function EstBundlePartage(sCode : string) : boolean;
  var sBaseMs,UneBase : string;
  begin
    result := (sCode = sCodeRDoss);
    if not result then
    begin
      sBaseMs := sBases;
      UneBase := ReadTokenSt (sBaseMs) ;
      while (UneBase <> '') and (not result) do
      begin
        result := not ExisteSql('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="YEB" '
              + 'AND CO_LIBRE="' + sCode + '" AND CO_ABREGE<>"BUN" AND CO_LIBELLE NOT IN'
              + '(SELECT IIF(DS_VUE="",DS_NOMTABLE,DS_NOMTABLE||";"||DS_VUE) AS CO_LIBELLE '
              + 'FROM ' + GetBase (UneBase, 'DESHARE')+')');
        UneBase := ReadTokenSt (sBaseMs) ;
      end;
    end;
  end;

  procedure TobBundleAddChampSup(t : tob);
  begin
    if assigned (t) then
    begin
      t.AddChampSupValeur ('PARTAGE', false, false);
      t.AddChampSupValeur ('PARTAGESTR', '-', false);
      t.AddChampSupValeur ('MODIFIE', false, false);
      t.AddChampSup ('IM_PARTAGE', false);
      t.AddChampSup ('IM_MODIFIE', false);
    end;
  end;

begin
  TobBundle.LoadDetailDBFromSQL ('COMMUN',
     'SELECT CO_CODE, CO_LIBELLE, CO_ABREGE FROM COMMUN WHERE CO_TYPE="YBU" ORDER BY CO_ABREGE');

  for iInd := 0 to TobBundle.Detail.Count -1 do
  begin
    TobB := TobBundle.Detail[iInd];
    TobBundleAddChampSup(TobB);

    if EstBundlePartage(TobB.GetString ('CO_CODE')) then
      partage(TobB,false)
    else departage(TobB,false);

    { on charge les bundles dépendants }
    TobB.LoadDetailDBFromSql ('COMMUN', 'SELECT C1.CO_LIBELLE AS CO_CODE, (SELECT C2.CO_LIBELLE FROM COMMUN C2 WHERE C2.CO_TYPE="YBU" AND C2.CO_CODE=C1.CO_LIBELLE) AS CO_LIBELLE FROM COMMUN C1 WHERE C1.CO_TYPE="YEB" AND C1.CO_LIBRE="' + TobB.GetString('CO_CODE') + '" AND C1.CO_ABREGE="BUN"');
  end;
end;

{ le bundle passé en paramètre dépend .. }
procedure TOF_YYBUNDLE.ChargeTobTVSlave(TobB : tob);
var sLib : string;
begin
  TobTVSlave.ClearDetail;
  if assigned(TobB) then
  begin
    TobTVSlave.SetString('CODE', TobB.GetString ('CO_CODE'));
    TobTVSlave.SetString('LIBELLE', TobB.GetString ('CO_LIBELLE'));
    if not CreeTobNoeuds (TobTVSlave, '', '') then
    begin
      sLib := TobBundle.FindFirst(['CO_CODE'], [stCircularite], false).GetString('CO_LIBELLE');
      PgiBox ('Réference de bundle Circulaire : ' + stCircularite);
      cbMultiDossier.Enabled := false;
    end;
  end;

  if TobTVSlave.Detail.Count <= 0 then
  begin
    TobTVSlave.SetString('CODE', '');
    TobTVSlave.SetString('LIBELLE', '< Pas de dépendances >');
  end;
end;

function TOF_YYBUNDLE.GetBasesRegroupement : string;
begin
  result := sBasesRegroupement;
end;

procedure TOF_YYBUNDLE.SetBasesRegroupement(LesBases : string);
begin
  sBasesRegroupement := LesBases;
end;

function TOF_YYBUNDLE.GetAction : tactionfiche;
begin
  result := action;
end;

procedure TOF_YYBUNDLE.SetAction(typeaction : tactionfiche);
begin
  action := typeaction;
end;

{ on passe une tob bundle en paramètre }
procedure TOF_YYBUNDLE.ChargeTobTVMaster(TobB : tob);
var sLib : string;
begin
  TobTVMaster.ClearDetail;
  if assigned(TobB) then
  begin
    TobTVMaster.SetString('CODE', TobB.GetString ('CO_CODE'));
    TobTVMaster.SetString('LIBELLE', TobB.GetString ('CO_LIBELLE'));
    if not CreeTobNoeudm (TobTVMaster, '', '') then
    begin
      sLib := TobBundle.FindFirst(['CO_CODE'], [stCircularite], false).GetString('CO_LIBELLE');
      PgiBox ('Réference de bundle Circulaire : ' + stCircularite);
      cbMultiDossier.Enabled := false;
    end;
  end;

  if TobTVMaster.Detail.Count <= 0 then
  begin
    TobTVMaster.SetString('CODE', '');
    TobTVMaster.SetString('LIBELLE', '< Pas de dépendances >');
  end;
end;

procedure tValideBundle.ValideParametrage;
var
  TobRegroup : TOB;
begin
  Inherited ;
  wInitProgressForm(nil, TraduireMemoire('Enregistrement du paramétrage'), TraduireMemoire('Partage du référentiel'), 6, False, True);
  { on supprime l'ancien partage }
  DeleteDeShare;

  if V_PGI.IoError = oeOk then
  begin
    wMoveProgressForm ('Sauvegarde des regroupements');

    ExecuteSql ('DELETE FROM YMULTIDOSSIER WHERE YMD_CODE="' + MS_CODEREGROUPEMENT + '"');

    if (V_PGI.IoError = oeOk) then //and (TobShare.Detail.Count > 0) then
    begin
      TobRegroup := Tob.Create ('', nil, -1);
      TobRegroup.LoadDetailDBFromSQL ('YMULTIDOSSIER', 'SELECT * FROM YMULTIDOSSIER WHERE YMD_CODE="' + FCodeRegroupement + '"');
      if TobRegroup.detail.Count > 0 then
      begin
        TobRegroup.Detail[0].SetString ('YMD_LIBELLE', 'Multi sociétés');
        TobRegroup.Detail[0].SetString ('YMD_CODE', MS_CODEREGROUPEMENT);
        TobRegroup.Detail[0].SetAllModifie(True);
        TobRegroup.InsertDB (nil);
      end;
      TobRegroup.Free;
    end;
  end;

  if V_PGI.IoError = oeOk then
  begin
    wMoveProgressForm ('Enregistrement dans la base commune');
    EnregistreParametrage (true);
  end;

  // EpureBase doit être fait en premier et pour cela plus de paramétrage dans les bases secondaires
  if V_PGI.IoError = oeOk then
  begin
    if FSuppData then
    begin
      wMoveProgressForm ('Epuration des tables partagées');
      EpureBase;
    end;
  end;

  if V_PGI.IoError = oeOk then
  begin
    wMoveProgressForm ('Enregistrement dans la(les) base(s) dépendante(s)');
    EnregistreParametrage (false);
  end;

  if V_PGI.IoError = oeOk then
  begin
    if FMajView then
    begin
        wMoveProgressForm ('Mise à jour des vues de la base commune ' + V_PGI.FCurrentAlias);
        PutToDebuglog ('Mise à jour des vues de la base commune ' +  V_PGI.FCurrentAlias +' à '+TimeToStr(Now),True);
      MajView;
    end;
  end;

  if V_PGI.IoError = oeOk then
  begin
    if ExecuteSQL ('DELETE FROM YDATATYPELINKS') = -1 then
      V_PGI.IoError := oeUnknown
    else if ExecuteSQL ('DELETE FROM YDATATYPETREES') = -1 then
      V_PGI.IoError := oeUnknown
    else if ExecuteSql ('UPDATE PARAMSOC SET SOC_DATA="-" WHERE (SOC_NOM="SO_RTACTTABHIE" OR SOC_NOM="SO_RTPROPTABHIE")') = -1 then
      V_PGI.IoError := oeUnknown;
  end;
  wFiniProgressForm;  
end ;

procedure TOF_YYBUNDLE.partage(t : tob ; withModifie : boolean);
begin
  t.SetBoolean('PARTAGE',true);
  t.SetString('PARTAGESTR','X');
  t.SetString('IM_PARTAGE','#ICO#'+intToStr(GetImage(imPartage)));
  if withModifie then
  begin
    t.SetBoolean('MODIFIE',true);
    t.SetString ('IM_MODIFIE', '#ICO#'+intToStr(imModifie));
  end;
end;

procedure TOF_YYBUNDLE.departage(t : tob ; withModifie : boolean);
begin
  t.SetBoolean('PARTAGE',false);
  t.SetString('PARTAGESTR','-');
  t.SetString('IM_PARTAGE','#ICO#'+intToStr(GetImage(imNonPartage)));
  if withModifie then
  begin
    t.SetBoolean('MODIFIE',true);
    t.SetString ('IM_MODIFIE', '#ICO#'+intToStr(imModifie));
  end;
end;

function TOF_YYBUNDLE.GetImage(image  : integer) : integer;
begin
  //if GetAction = taconsult then
  // result := imNonDefini
  //else
  result := image;
end;

procedure TOF_YYBUNDLE.ChargeLesBitmaps;
begin
  ChargeImageList;
  if assigned(GetControl('MNAJOUTE')) then
    V_PGI.GraphList.GetBitmap(52-1, TMenuItem(GetControl('MNAJOUTE')).Bitmap);
  if assigned(GetControl('MNRETIRE')) then
    V_PGI.GraphList.GetBitmap(53-1, TMenuItem(GetControl('MNRETIRE')).Bitmap);
  if assigned(GetControl('MNDETAIL')) then
    V_PGI.GraphList.GetBitmap(94-1, TMenuItem(GetControl('MNDETAIL')).Bitmap);
end;

procedure TOF_YYBUNDLE.OnArgument (S : String ) ;

  procedure InitControl;
  begin
    if assigned(GetControl('GBUNDLE')) then
    begin
      GBundle := THGrid(GetControl('GBUNDLE'));
      GBundle.OnRowEnter := BundleOnRowEnter;
      //GBUNDLE.ColAligns [2] := taCenter;
    end;

    { boutons }
    if assigned(GetControl('TBBACTUVIEW')) then
    begin
      TBBActuView := TToolBarButton97(GetControl('TBBACTUVIEW'));
      TBBActuView.OnClick := TBBActuViewOnClick;
    end;

    if assigned(GetControl('TBBEPURTABLE')) then
    begin
      TBBEpurTable := TToolBarButton97(GetControl ('TBBEPURTABLE'));
      TBBEpurTable.OnClick := TBBEPURTABLEOnClick;
    end;

    if assigned(GetControl ('TBBEXPORT')) then
    begin
      TToolBarButton97(GetControl ('TBBEXPORT')).OnClick := TBBExportOnClick;
    end;

    if assigned(GetControl ('TBBIMPORT')) then
    begin
      TToolBarButton97(GetControl ('TBBIMPORT')).OnClick := TBBImportOnClick;
    end;

    if assigned(GetControl ('MULTIDOSSIER')) then
    begin
      CbMultiDossier := THValComboBox (GetControl ('MULTIDOSSIER'));
      CbMultiDossier.OnClick := CbMultiDossierOnClick;
    end;

    if Assigned (GetControl ('TVSLAVE')) then
    begin
      TVSlave := TTreeView(GetControl ('TVSLAVE'));
      ChargeImageList;
      TVSlave.Images := V_PGI.GraphList;
    end;

    if Assigned (GetControl ('TVMASTER')) then
    begin
      TVMaster := TTreeView(GetControl ('TVMASTER'));
      ChargeImageList;
      TVMaster.Images := V_PGI.GraphList;
    end;

    if assigned(GetControl('BUNDLEAFFICHE')) then
      ThValComboBox(GetControl('BUNDLEAFFICHE')).OnChange := BundleAfficheOnChange;

    if assigned(GetControl('SPLITVERT')) then
      ThSplitter(GetControl('SPLITVERT')).OnMoved := SplitMoved;

    if assigned(GetControl('MNAJOUTE')) then
      TMenuItem(GetControl('MNAJOUTE')).OnClick := MnAjouteBundleClick;

    if assigned(GetControl('MNRETIRE')) then
      TMenuItem(GetControl('MNRETIRE')).OnClick := MnRetireBundleClick;

    if assigned(GetControl('MNDETAIL')) then
      TMenuItem(GetControl('MNDETAIL')).OnClick := MnDetailBundleClick;

    if assigned(GetControl('MENUPOP')) then
      TPopupMenu(GetControl('POPACTION')).OnPopup := PopActionOnPopup;

    if assigned(GetControl ('BPARAM')) then
    begin
      TToolBarButton97(GetControl ('BPARAM')).OnClick := bParamClick;
    end;

    if assigned(GetControl ('BVOIR')) then
    begin
      TToolBarButton97(GetControl ('BVOIR')).OnClick := BVoirOnClick;
    end;

    if assigned(GetControl('MNH')) then
      TMenuItem(GetControl('MNH')).OnClick := mnhOnClick;

    if assigned(GetControl('MNB')) then
      TMenuItem(GetControl('MNB')).OnClick := mnbOnClick;

  end;

  procedure createTob;
  begin
    TobBundle  := Tob.Create ('COMMUN', nil, -1);
    { pour gérer l'affichage }
    TobGBundle := Tob.Create ('', nil, -1);
    TobTVSlave     := Tob.Create ('', nil, -1);
    TobTVMaster     := Tob.Create ('', nil, -1);
  end;

begin
  Inherited ;

  TForm(Ecran).Caption := TraduireMemoire('Partage du référentiel');
  UpdateCaption(Ecran);

  { tobs }
  createTob;
  TobTVSlave.AddChampSup ('CODE', false);
  TobTVSlave.AddChampSup ('LIBELLE', false);
  TobTVMaster.AddChampSup ('CODE', false);
  TobTVMaster.AddChampSup ('LIBELLE', false);

  { controls }
  initControl;
  RedimensionnePanelsTV;
  ChargeLesBitmaps;

  { init des variables }
  bRegroupementOk := false;
  bModifie := false;
  bBasePrincipale := false;
  SetAction(tamodif);
  ValideBundle := tValideBundle.Create('',TobBundle);
end ;

procedure TOF_YYBUNDLE.OnClose ;

  procedure freeTob;
  begin
    TobBundle.Free;
    TobGBundle.Free;
    TobTVSlave.Free;
    TobTVMaster.Free;
  end;

begin
  if (bModifie) and bBasePrincipale then
  begin
    if PgiAsk (TraduireMemoire('Voulez vous enregistrer les modifications ?'), '') = mrYes then
      OnUpdate
  end;
  Inherited ;
    freeTob;
    ValideBundle.Free;
end;

function TOF_YYBUNDLE.OkConnect(S:String) : boolean;
var UneBase:String;
    Ok        : Boolean ;
    DBSource: TDatabase;
begin
  result:=true;

  UneBase := ReadTokenSt (S) ;
  while UneBase <> '' do
  begin
    Application.ProcessMessages ;
    DBSource := TDatabase.Create(nil);
    ok := ConnecteDB('',DBSource,UneBase);
    if not ok then
    begin
      PGIBOX(TraduireMemoire('Erreur lors de la connexion à la base '+ UneBase +' : validation impossible !' +
      chr(13) + TraduireMemoire(' Veuillez vérifier vos éléments de connexion (CEGIDPGI.INI ...)')));
      result:=false;
      abort; // base inaccessible
    end;
    UneBase := ReadTokenSt (S) ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JSI
Créé le ...... : 18/04/2006
Modifié le ... :   /  /    
Description .. : Mise à jour de la table deshare
Suite ........ : si modification du contenu d'un bundle :
...............  si bDeleteData : les tables ajoutées seront vidées
Mots clefs ... : BUNDLE;DESHARE;
*****************************************************************}
procedure MajBundlePartage(bDeleteData : boolean = false);
var
    sAbrege : string;
    sTableVue, sTable, sVue : string;
    sBasePrincipale : string;
    bModif : boolean;
    tobBundle, tobElementBundle : tob;
    iInd, iInd2 : integer;

  function EstBundlePartage(sBundle : string) : boolean;
  begin
    { le bundle est partagé si au moins un de ses éléments est présent dans la table deshare }
    result := ExisteSql('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="YEB" '
                + 'AND CO_LIBRE="' + sBundle + '" AND CO_ABREGE<>"BUN" '
                + 'AND CO_CODE NOT IN'
                 { on exclut les éléments suceptibles d'être dans plusieurs bundles }
                + '(SELECT C1.CO_CODE FROM COMMUN C1, COMMUN C2 '
                + 'WHERE C1.CO_TYPE="YEB" AND C2.CO_TYPE="YEB" AND C1.CO_LIBELLE=C2.CO_LIBELLE '
                + 'AND C1.CO_LIBRE<>C2.CO_LIBRE AND C1.CO_ABREGE<>"BUN" AND C2.CO_ABREGE<>"BUN") '
                 { on teste la présence de ces éléments dans la table deshare }
                + 'AND CO_LIBELLE IN'
                + '(SELECT IIF(DS_VUE="",DS_NOMTABLE,DS_NOMTABLE||";"||DS_VUE) AS ELEMENT '
                + 'FROM DESHARE)');
  end;

  function getBasePrincipale() : string;
  var Qbase : tQuery;
  begin
    result := V_PGI.SchemaName;
    Qbase := openSql('SELECT ##TOP 1## DS_NOMBASE FROM DESHARE',true);
    try
      if not Qbase.Eof then
        result := Qbase.FindField('DS_NOMBASE').asString;
    finally
      ferme(Qbase);
    end;
  end;

  procedure deleteData(laTable : string);
  begin
    if sBasePrincipale <> V_PGI.SchemaName then
    begin
      { le @@ pour ne pas traduire la requete et donc ne pas utiliser DESHARE }
      ExecuteSql ('@@DELETE FROM ' + laTable);
      ExecuteSql ('@@DELETE FROM CHOIXCOD WHERE CC_TYPE IN (SELECT DO_TYPE FROM DECOMBOS WHERE ' +
                  'EXISTS (SELECT 1 FROM DESHARE WHERE DS_TYPTABLE="TTE" AND DS_NOMTABLE=DO_COMBO))');
    end;
  end;

  procedure MajView();
  var  bSaveEnableDeShare, bSaveExisteDeShare : boolean;
  begin
    if sBasePrincipale = V_PGI.SchemaName then
    begin
      DBDeleteAllView(DBSOC, Nil, V_PGI.Driver);
      DBCreateAllView(DBSOC, Nil, V_PGI.Driver);
    end else
    begin
      bSaveEnableDeShare := V_PGI.enableDEShare;
      bSaveExisteDeShare := V_PGI.existeDEShare;

      ChargeDeShare ();

      V_PGI.enableDEShare := true;
      V_PGI.existeDEShare := true;

      MajOracleSession(DBSOC);

      DBDeleteAllView(DBSOC, Nil, V_PGI.Driver);
      DBCreateAllView(DBSOC, Nil, V_PGI.Driver);

      V_PGI.enableDEShare := bSaveEnableDeShare;
      V_PGI.existeDEShare := bSaveExisteDeShare;
    end;
  end;

begin
  bModif := false;
  tobBundle := tob.Create('',nil,-1);
  try
    { on récupère le nom de la base principale }
    tobBundle.LoadDetailFromSQL('SELECT * FROM COMMUN WHERE CO_TYPE="YBU"');
    sBasePrincipale := getBasePrincipale();

    for iInd := 0 to tobBundle.Detail.Count -1 do
    begin
      { on teste si le bundle est partagé }
      if EstBundlePartage(tobBundle.Detail[iInd].GetString('CO_CODE')) then
      begin
        tobElementBundle := tob.Create('',nil,-1);
        try
          { on cherche les infos concernant le nouvel élément }
          tobElementBundle.LoadDetailFromSQL('SELECT * FROM COMMUN WHERE CO_TYPE="YEB" '
                    + 'AND CO_LIBRE="' + tobBundle.Detail[iInd].GetString('CO_CODE') + '" '
                    + 'AND CO_ABREGE<>"BUN" AND CO_LIBELLE NOT IN'
                    + '(SELECT IIF(DS_VUE="",DS_NOMTABLE,DS_NOMTABLE||";"||DS_VUE) AS ELEMENT '
                    + 'FROM DESHARE)');
          for iInd2 := 0 to tobElementBundle.Detail.Count -1 do
          begin
            sTableVue := tobElementBundle.Detail[iInd2].GetString('CO_LIBELLE');
            sAbrege   := tobElementBundle.Detail[iInd2].GetString('CO_ABREGE');
            { test de l'association d'une table et d'une vue }
            if (sAbrege = 'TAB') and (pos (';', sTableVue) > 0) then
            begin
              sTable := readTokenSt(sTableVue);
              sVue   := readTokenSt(sTableVue);
            end else
            begin
              sTable := sTableVue;
              sVue   := '';
            end;

            { on vérifie l'existence au cas où }
            if not ExisteSQL('SELECT DS_NOMTABLE FROM DESHARE '
                           + 'WHERE DS_NOMTABLE = "' + sTable + '" '
                           + 'AND DS_MODEFONC = "LIB"') then
            begin
              { on insère dans la base partagée }
              executeSql('INSERT INTO DESHARE '
                       + '(DS_MODEFONC,DS_NOMBASE,DS_NOMTABLE,DS_TYPTABLE,DS_VUE) '
                       + 'VALUES ("LIB","' + sBasePrincipale + '","' + sTable + '"'
                       + ',"' + sAbrege + '","' + sVue + '")');

              { on supprime les données de la table }
              if bDeleteData then
                if sAbrege= 'TAB' then
                  deleteData(sTable);

              { flag la modification pour la maj des vues }
              bModif := true;
            end;
          end;
        finally
          tobElementBundle.Free;
        end;
      end;
    end;
    { maj des vues }
    if bModif then MajView;
  finally
    tobBundle.Free;
  end;
end;

function CreerVueMultiSoc (NomVue, stSql : string) : boolean;
var
  bSaveEnableDeShare, bSaveExisteDeShare : boolean;
  LeSql : string;

  function ExisteVueEtRecupereSql : boolean;
  var
    Q : TQuery;
  begin
    Q := OpenSQL ('SELECT DV_SQL FROM DEVUES WHERE DV_NOMVUE="' + NomVue + '"', true);
    if not Q.Eof then
    begin
      Result := true;
      LeSql := Q.Fields [0].AsString
    end else
    begin
      Result := false;
      PgiBox ('La vue n''existe pas');
    end;
    Ferme (Q);
  end;

begin
  Result := false;
  if EstMultiSoc then
  begin
    if ExisteVueEtRecupereSql then
    begin
      if stSql = '' then stSql := LeSql;
      
      bSaveEnableDeShare := V_PGI.enableDEShare;
      bSaveExisteDeShare := V_PGI.existeDEShare;

      V_PGI.enableDEShare := true;
      V_PGI.existeDEShare := true;

      DBDeleteView (DBSOC, V_PGI.Driver, NomVue);
      Result := DBCreateView (DBSOC, NomVue, stSql, V_PGI.Driver);

      V_PGI.enableDEShare := bSaveEnableDeShare;
      V_PGI.existeDEShare := bSaveExisteDeShare;
    end;
  end;
end;

procedure TOF_YYMAJVUEMULTISOC.OnArgument (S : String ) ;
begin
  TToolBarButton97(GetControl('BVALIDER')).OnClick := bValiderClick;
end;

procedure TOF_YYMAJVUEMULTISOC.BValiderClick(Sender: TObject);
begin
  if CreerVueMultiSoc (GetControlText ('LAVUE'), GetControlText ('MSQL')) then
    PgiBox ('La vue ' + GetControlText ('LAVUE') + ' a été créée avec succès')
  else
    PgiBox ('La vue ' + GetControlText ('LAVUE') + ' n''a pas pu être créée');
end;

Initialization
  registerclasses ( [ TOF_YYBUNDLE ] ) ;
  registerclasses ( [ TOF_YYMAJVUEMULTISOC ] ) ;
end.
