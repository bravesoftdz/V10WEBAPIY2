{***********UNITE*************************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 20/09/2001
Modifié le ... :   /  /
Description .. : MUL permettant de consulter le stock sur un site distant qui
Suite ........ : possède un serveur eAGL.
Suite ........ : Ce MUL fonctionne aussi bien en local que sur site distant.
Mots clefs ... : STOCK,DISPO,EAGL,MUL
*****************************************************************}
unit UtofDispoDist_Mul;

interface

uses UTOF, HCtrls, Windows, uHttp, HDimension, UTOB,
  {$IFDEF EAGLCLIENT}
  eMul, eFiche,
  {$ELSE}
  Mul, Fiche, uHttpCS, RasTool, PARAMSITEDIST_TOF, Ras,
  dbTables, LicUtil, Fe_Main,
  {$ENDIF}
  SysUtils, Controls, EntGC, UtilGC, HMsgBox, HEnt1, Forms, StdCtrls,
  ParamSoc, HTB97, Classes, M3FP;

type
  TOF_DISPODIST_MUL = class(TOF)
  private
    BMulDistant: boolean;
    StFieldList: string;
    {$IFNDEF EAGLCLIENT}
    GMul: THGrid;
    StServerAdd, StDossier, StUser, StPassword: string;
    StModeConnexion, RasEntry, RasUser, RasPass, RasTel: string;
    RasSavePass, IsConnected: boolean;
    hRas: THandle;
    aServer: THttpServer;
    function LanceConnexion(BDemarrer: boolean): boolean;
    procedure LanceSiteDistant;
    procedure AffichageMUL;
    {$ENDIF}
    procedure AffichageTitresColonnes(BTous: boolean);
    procedure DimensionsArticle(StArt: string; QuelCode: Integer);
    procedure LimiteLesCombosDimensions(TabDim, TabValDim: array of string);
    procedure MULSiteDistant(BDistant: Boolean);
  public
    ChangeListe: boolean;
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnClose; override;
  end;

implementation

uses
  FactUtil, Ent1;

{ TOF_DISPODIST_MUL }

procedure TOF_DISPODIST_MUL.OnArgument(Arguments: string);
var THLIB: THLabel;
  iCol, Nbr: integer;
  FF: TFMUL;
  BVisible: Boolean;
  Ctrl: TControl;
  Etab, LstDepot, Stg: string;
begin
  inherited;
  FF := TFMul(Ecran);
  TFMul(Ecran).OldManuel := True;
  ChangeListe := False;
  {$IFNDEF EAGLCLIENT}
  StServerAdd := '';
  StDossier := '';
  StUser := '';
  StPassword := '';
  GMul := THGrid(GetControl('GCONSULT'));
  {$ENDIF}

  // Mise en forme des libellés des dimensions    JD : 20/09/2002
  for iCol := 1 to MaxDimension do
  begin
    THLIB := THLabel(GetControl('DIMENSION' + IntToStr(iCol)));
    THLIB.Caption := RechDom('GCCATEGORIEDIM', 'DI' + InttoStr(iCol), False);
    BVisible := not ((Copy(THLIB.Caption, 1, 2) = '.-') or (THLIB.Caption = '??'));
    THLIB.Visible := BVisible;
    TControl(GetControl(THLIB.FocusControl.Name)).Visible := BVisible;
  end;

  // Paramètrage des libellés des onglets des zones et tables libres du dépôt
  if not VH_GC.GCMultiDepots then
  begin
    SetControlCaption('PTABLESLIBRESDEP', 'Tables libres étab.');
    SetControlCaption('PZONESLIBRESDEP', 'Zones libres étab.');
  end;

  // Paramétrage des libellés des tables libres articles et dépôts
  Nbr := 0;
  if (GCMAJChampLibre(FF, False, 'EDIT', 'GDE_LIBREDEP', 10, '') = 0) then SetControlVisible('PTABLESLIBRESDEP', False);
  // Mise en forme des libellés des dates, booléans libres et montants libres
  if (GCMAJChampLibre(FF, False, 'EDIT', 'GDE_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VALDEP', False) else inc(Nbr);
  if (GCMAJChampLibre(FF, False, 'EDIT', 'GDE_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATEDEP', False) else inc(Nbr);
  if (GCMAJChampLibre(FF, False, 'BOOL', 'GDE_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOLDEP', False) else inc(Nbr);
  if (Nbr = 0) then SetControlVisible('PZONESLIBRESDEP', False);

  SetControlText('XX_WHERE_STOCK', 'GQ_CLOTURE<>"X" and GQ_PHYSIQUE>0');

  if Arguments = 'MULDISTANT' then
  begin
    {$IFNDEF EAGLCLIENT}
    MULSiteDistant(True);
    AffichageTitresColonnes(True);
    {$ENDIF}
  end else
  begin
    if Arguments = 'MULLOCAL' then SetControlVisible('BDISTANT', False);
    MULSiteDistant(False);
  end;

  // ne proposer que les dépôts de l'établissement de l'utilisateur
  Etab := EtabForce;
  if Etab <> '' then
  begin
    if GetParamSoc('SO_GCMULTIDEPOTS') then
      LstDepot := ListeDepotParEtablissement(Etab)
    else
      LstDepot := '"'+ Etab +'"';
    if LstDepot <> '' then
    begin
      Ctrl := GetControl('GQ_DEPOT');
      if (Ctrl <> nil) and (Ctrl is THMultiValComboBox) then
      begin
        Stg := THMultiValComboBox(Ctrl).Plus;
        if Stg <> '' then Stg := Stg + ' AND ';
        Stg := Stg + 'GDE_DEPOT IN ('+ LstDepot +')';
        THMultiValComboBox(Ctrl).Plus := Stg;
        THMultiValComboBox(Ctrl).Complete := True;
      end;
    end;
  end;

  SetFocusControl('_GA_CODEBARRE');
end;

procedure TOF_DISPODIST_MUL.OnLoad;
begin
  inherited;
  NextControl(Ecran);
  if (GetControlText('GA_CODEARTICLE') = '') then
  begin
    if (GetControlText('_GA_CODEBARRE') = '') then
    begin
      PGIInfo(TraduireMemoire('Vous devez choisir un article avant de lancer la consultation du stock.'), ecran.caption);
      TFMul(ECRAN).Q.Manuel := True;
      exit;
    end;
  end;
  if not BMulDistant then
    TFMul(ECRAN).Q.Manuel := False;
end;

procedure TOF_DISPODIST_MUL.OnUpdate;
var xx_where: string;
begin
  inherited;
  // Gestion des checkBox : booléens libres  dépôts
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'BOOL', 'GDE_BOOLLIBRE', 3, '');
  // Gestion des dates libres dépôts
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'DATE', 'GDE_DATELIBRE', 3, '_');
  // Gestion des montants libres dépôts
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'EDIT', 'GDE_VALLIBRE', 3, '_');
  SetControlText('XX_WHERE', xx_where);
  AffichageTitresColonnes(True);
  if ((GetControlText('GA_CODEARTICLE') = '') and (GetControlText('_GA_CODEBARRE') = '')) then exit;
  {$IFNDEF EAGLCLIENT}
  if BMulDistant then
  begin
    if not IsConnected then
      if not LanceConnexion(True) then exit;
    AffichageMUL;
  end;
  {$ENDIF}
end;

procedure TOF_DISPODIST_MUL.AffichageTitresColonnes(BTous: boolean);
var F: TFMUL;
  TLIB: THLabel;
  TOBVide: TOB;
  Col, Dec, ColEAGL: Integer;
  NomCol, Perso, NomList, StA, St, FF: string;
  FRecordSource, FLien, FFieldList, FTitre, FSortBy, FLargeur, FAlignement, FParams, tt, NC: string;
  OkTri, OkNumCol, Obli, OkLib, Sep, OkVisu, OkNulle, OkCumul: boolean;
begin
  F := TFMul(Ecran);
  NomList := F.Q.Liste;
  // Mise en place des libellés des dimensions dans les colonnes
  {$IFNDEF EAGLCLIENT}
  if BTous then
  begin
    TOBVide := TOB.Create('aucune table', nil, -1);
    TOBVide.PutGridDetailOnListe(GMul, F.Q.Liste);
    F.Hmtrad.ResizeGridColumns(GMul);
    TOBVide.Free;
  end;
  {$ENDIF}
  Col := 0;
  ColEAGL := 0;
  StFieldList := '';
  ChargeHListe(NomList, FRecordSource, FLien, FSortBy, FFieldList, FTitre, FLargeur, FAlignement, FParams, tt, NC, Perso, OkTri, OkNumCol);
  while Ftitre <> '' do
  begin
    StA := ReadTokenSt(FAlignement);
    St := ReadTokenSt(Ftitre);
    Nomcol := ReadTokenSt(FFieldList);
    if NomCol <> '' then
    begin
      TransAlign(StA, FF, Dec, Sep, Obli, OkLib, OkVisu, OkNulle, OkCumul);
      if OkVisu then
      begin
        if (copy(St, 1, 3) = 'DIM') then
        begin
          TLIB := THLabel(GetControl('DIMENSION' + copy(St, 5, 1)));
          {$IFNDEF EAGLCLIENT}
          GMul.Cells[Col, 0] := TLIB.Caption;
          {$ENDIF}
          if not TFMul(Ecran).Q.Manuel then
          {$IFDEF EAGLCLIENT}
            F.FListe.Cells[ColEAGL, 0] := TLIB.Caption;
          {$ELSE}
            F.Fliste.columns[Col].Field.DisplayLabel := TLIB.Caption;
          {$ENDIF}
        end;
        if StFieldList = '' then StFieldList := Nomcol else StFieldList := StFieldList + ';' + Nomcol;
        inc(Col);
      end;
      Inc(ColEAGL);
    end;
  end;
  ChangeListe := False;
end;

{$IFNDEF EAGLCLIENT}

procedure TOF_DISPODIST_MUL.AffichageMUL;
var MaTOB: TOB;
  StSelect, StInifile, MsgServeur: string;
begin
  StInifile := 'CEGIDPGI.ini';
  aServer := ConnectHttpServer(StServerAdd, MsgServeur);
  if MsgServeur <> '' then
  begin
    PGIInfo(TraduireMemoire(MsgServeur), ecran.caption);
    exit;
  end;
  if (aServer <> nil) and (aServer.IsConnected) then
  begin
    if not ConnecteAGLServer(aServer, StUser, StPassword, StDossier, StInifile) then
    begin
      PGIInfo(TraduireMemoire('Impossible de se connecter à la base distante, vérifier les paramètres du site distant.'), ecran.caption);
      exit;
    end;
  end;
  StSelect := TRim(TFMul(Ecran).Q.SQL[0]);
  MaTOB := nil;
  GMul.VidePile(False);
  try
    StSelect := StringReplace(StSelect, '''', '"', [rfReplaceAll]);
    MaTOB := OpenSQLeAGLServer(aServer, StSelect, True);
    if MaTOB <> nil then
    begin
      MaTOB.PutGridDetail(GMul, False, False, StFieldList);
      TFMul(Ecran).HMTrad.ResizeGridColumns(GMul);
    end
    else PGIInfo(TraduireMemoire('Impossible de récupérer des données sur le serveur distant.'), ecran.caption);
  finally
    if MaTOB <> nil then MaTOB.Free;
    DeconnecteAGLServer(aServer);
    DeconnectHttpServer(aServer);
  end;
end;

procedure TOF_DISPODIST_MUL.LanceSiteDistant;
begin
  AglRastool(hRas, RasEntry, RasUser, RasPass, RasSavePass);
end;

function TOF_DISPODIST_MUL.LanceConnexion(BDemarrer: boolean): boolean;
var BConnexion: boolean;
  Reponse: DWORD;
begin
  result := False;
  SourisSablier;
  if (StModeConnexion = '') then
  begin
    ChargeParamSiteDistant(StServerAdd, StDossier, StUser, StPassword, StModeConnexion,
      RasEntry, RasUser, RasPass, RasTel);
    if (StModeConnexion = '') and (BDemarrer) then
    begin
      PGIError('Votre mode de connexion au site distant n''est pas correctement paramétré !', ecran.caption);
      SourisNormale;
      exit;
    end;
  end;
  if StModeConnexion = '002' then
  begin
    BConnexion := AglGetActiveRAS(RasEntry, hRas);
    if BDemarrer then
    begin
      if not BConnexion then
      begin
        Reponse := AglInitRasDial(RasEntry, RasTel, RasUser, RasPass, hRas, nil);
        if (RasEntry = '') or (Reponse <> 0) then
        begin
          PGIError('Votre accès distant n''est pas correctement paramétré !', ecran.caption);
          SourisNormale;
          exit;
        end;
      end;
      IsConnected := True;
    end
    else
    begin
      if BConnexion then
      begin
        Reponse := AglRasHangup(hRas);
        if Reponse > 600 then
        begin
          PGIInfo('La connexion n''a pas été fermée correctement, veuillez la fermer manuellement.', ecran.caption);
          SourisNormale;
          exit;
        end;
      end;
      IsConnected := False;
    end;
  end
  else IsConnected := BDemarrer;
  result := True;
  SetControlVisible('BCONNECT', not BDemarrer);
  SetControlVisible('BDECONNECT', BDemarrer);
  SourisNormale;
end;
{$ENDIF}

procedure TOF_DISPODIST_MUL.OnClose;
begin
  inherited;
  {$IFNDEF EAGLCLIENT}
  GMul.VidePile(False);
  LanceConnexion(False);
  {$ENDIF}
end;

procedure TOF_DISPODIST_MUL.DimensionsArticle(StArt: string; QuelCode: Integer);
var Q: TQuery;
  i: integer;
  StCodeArticle, StCodeBarre: string;
  TabDim: array[1..MaxDimension] of string;
  TabValDim: array[1..MaxDimension] of string;
begin
  //QuelCode=1 correspond à GA_ARTICLE et QuelCode=2 correspond à GA_CODEBARRE
  for i := 1 to MaxDimension do TabValDim[i] := '';
  for i := 1 to MaxDimension do TabDim[i] := '';
  StCodeArticle := '';
  if QuelCode = 1 then StCodeArticle := GetControlText('GA_CODEARTICLE')
  else if QuelCode = 2 then
  begin
    StCodeBarre := GetControlText('_GA_CODEBARRE');
    if StCodeBarre <> '' then
    begin
      SetControlText('GA_CODEARTICLE', '');
      Q := OpenSQL('SELECT GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEBARRE="' + StCodeBarre + '"',
        True);
      if not Q.Eof then
      begin
        for i := 1 to MaxDimension do TabValDim[i] := Q.FindField('GA_CODEDIM' + IntToStr(i)).AsString;
        StCodeArticle := Q.FindField('GA_CODEARTICLE').AsString;
        SetControlText('GA_CODEARTICLE', StCodeArticle);
      end
      else
        PGIInfo(TraduireMemoire('Impossible de trouver un article correspondant à ce code barre.'), ecran.caption);
      Ferme(Q);
    end;
  end;
  if (StCodeArticle <> '') then
  begin
    Q := OpenSQL('SELECT GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,GA_CODEBARRE FROM ARTICLE WHERE GA_CODEARTICLE="' + StCodeArticle
      + '" AND GA_STATUTART IN ("UNI","DIM")', True);
    if not Q.Eof then
    begin
      if QuelCode = 1 then
      begin
        if GetControlText('_GA_CODEBARRE') <> '' then
        begin
          if not ExisteSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEBARRE="' + GetControlText('_GA_CODEBARRE') + '" AND GA_CODEARTICLE="' +
            StCodeArticle + '"') then
            SetControlText('_GA_CODEBARRE', '');
        end;
      end;
      for i := 1 to MaxDimension do TabDim[i] := Q.FindField('GA_GRILLEDIM' + IntToStr(i)).AsString;
      LimiteLesCombosDimensions(TabDim, TabValDim);
      if not BMulDistant then SetControlText('XX_WHERE', '');
      SetControlCaption('LIBELLEARTICLE', RechDom('GCARTICLEGENERIQUE', StCodeArticle, False));
    end
    else
    begin
      PGIInfo(TraduireMemoire('Impossible de trouver cet article, vérifier son existence puis relancer la consultation.'), ecran.caption);
      TFMul(Ecran).Q.Manuel := True;
      SetControlCaption('LIBELLEARTICLE', '');
    end;
    Ferme(Q);
  end
  else
  begin
    if (GetControlText('_GA_CODEBARRE') = '') and (GetControlText('GA_CODEARTICLE') = '') then
    begin
      TFMul(Ecran).Q.Manuel := True;
      SetControlCaption('LIBELLEARTICLE', '');
    end;
  end;
end;

procedure TOF_DISPODIST_MUL.LimiteLesCombosDimensions(TabDim, TabValDim: array of string);
var i: integer;
  CBDim: THMultiValComboBox;
begin
  for i := 1 to MaxDimension do
  begin
    if i = 1 then CBDim := THMultiValComboBox(GetControl('GDI_LIBELLE'))
    else CBDim := THMultiValComboBox(GetControl('GDI_LIBELLE_' + IntToStr(i - 1)));
    if TabDim[i - 1] = '' then
    begin
      CBDim.Enabled := False;
      CBDim.Plus := '';
      CBDim.Value := TraduireMemoire('<<Tous>>');
      SetControlEnabled('DIMENSION' + IntToStr(i), False);
    end else
    begin
      CBDim.Enabled := True;
      SetControlEnabled('DIMENSION' + IntToStr(i), True);
      CBDim.Plus := 'AND GDI_GRILLEDIM="' + TabDim[i - 1] + '" ORDER BY GDI_RANG';
      if (TCheckBox(GetControl('CBRECHTAILLE')).Checked) and (TabValDim[i - 1] <> '') then
        CBDim.Value := RechDom('GCDIMENSION', TabValDim[i - 1], False, 'GDI_GRILLEDIM="' + TabDim[i - 1] + '"')
      else CBDim.Value := TraduireMemoire('<<Tous>>');
    end;
  end;
end;

procedure TOF_DISPODIST_MUL.MULSiteDistant(BDistant: Boolean);
var StPlus: string;
begin
  BMulDistant := BDistant;
  TFMul(Ecran).Q.Manuel := True;
  StPlus := '';
  if BDistant then
    Ecran.Caption := 'Consultation du stock sur site distant'
  else
  begin
    {$IFNDEF EAGLCLIENT}
    LanceConnexion(False);
    {$ENDIF}
    SetControlText('GA_CODEARTICLE', '');
    SetControlText('_GA_CODEBARRE', '');
    Ecran.Caption := 'Consultation du stock en local';
    if not GetParamSoc('SO_BTQPROXIMITE') then StPlus := 'AND GDE_SURSITE="X"';
  end;
  THMultiValComboBox(GetControl('GQ_DEPOT')).Plus := StPlus;
  UpdateCaption(Ecran);
  SetControlVisible('FListe', not BDistant);
  SetControlVisible('GConsult', BDistant);
  SetControlVisible('BCONNECT', BDistant);
  SetControlVisible('BDECONNECT', False);
  SetControlVisible('BPARAMCONNECT', BDistant);
  SetControlVisible('BLIGNERNIS', BDistant);
  TToolBarButton97(GetControl('BDISTANT')).Down := BDistant;
end;

///////////////////////////////////////////////////////////////////////////////

{$IFNDEF EAGLCLIENT}

procedure TOFConnexionSiteDistant(Parms: array of variant; nb: integer);
var F: TFMul;
  Latof: TOF;
begin
  F := TFMul(Integer(Parms[0]));
  if (F is TFMul) then Latof := TFMul(F).Latof else exit;
  if (Latof is TOF_DISPODIST_MUL) then
  begin
    if UpperCase(string(Parms[1])) = 'TRUE' then TOF_DISPODIST_MUL(Latof).LanceConnexion(True)
    else TOF_DISPODIST_MUL(Latof).LanceConnexion(False);
  end;
end;

procedure TOFConnexionRNIS(Parms: array of variant; nb: integer);
var F: TFMul;
  Latof: TOF;
begin
  F := TFMul(Integer(Parms[0]));
  if (F is TFMul) then Latof := TFMul(F).Latof else exit;
  if (Latof is TOF_DISPODIST_MUL) then TOF_DISPODIST_MUL(Latof).LanceSiteDistant;
end;

procedure TOFAfficheParamSiteDistant(Parms: array of variant; nb: integer);
var F: TFMul;
  Latof: TOF;
  LesParametres, UnParam, StParam, ValParam: string;
  x: integer;
begin
  LesParametres := '';
  F := TFMul(Integer(Parms[0]));
  if (F is TFMul) then Latof := TFMul(F).Latof else exit;
  if (Latof is TOF_DISPODIST_MUL) then
  begin
    LesParametres := AGLLanceFiche('MBO', 'PARAMSITEDIST', '', '', '');
    TToolbarButton97(TOF_DISPODIST_MUL(Latof).GetControl('BPARAMCONNECT')).Down := False;
    while LesParametres <> '' do
    begin
      UnParam := READTOKENST(LesParametres);
      if UnParam <> '' then
      begin
        x := pos('=', UnParam);
        if x <> 0 then
        begin
          StParam := copy(UnParam, 1, x - 1);
          ValParam := copy(UnParam, x + 1, length(UnParam));
          if StParam = 'SO_SDSERVERADDRESS' then TOF_DISPODIST_MUL(Latof).StServerAdd := ValParam
          else if StParam = 'SO_SDDOSSIER' then TOF_DISPODIST_MUL(Latof).StDossier := ValParam
          else if StParam = 'SO_SDUSERNAME' then TOF_DISPODIST_MUL(Latof).StUser := ValParam
          else if StParam = 'SO_SDPASSWORD' then TOF_DISPODIST_MUL(Latof).StPassword := ValParam
          else if StParam = 'SO_SDMODECONNEXION' then TOF_DISPODIST_MUL(Latof).StModeConnexion := ValParam
          else if StParam = 'SO_SDPROFIL' then TOF_DISPODIST_MUL(Latof).RasEntry := ValParam
          else if StParam = 'PROFILUSER' then TOF_DISPODIST_MUL(Latof).RasUser := ValParam
          else if StParam = 'PROFILPASS' then TOF_DISPODIST_MUL(Latof).RasPass := DeCryptageSt(ValParam)
          else if StParam = 'PROFILTEL' then TOF_DISPODIST_MUL(Latof).RasTel := ValParam
            ;
        end;
      end;
    end;
  end;
end;

procedure TOFMULDistant(Parms: array of variant; nb: integer);
var F: TFMul;
  Latof: TOF;
begin
  F := TFMul(Integer(Parms[0]));
  if (F is TFMul) then Latof := TFMul(F).Latof else exit;
  if (Latof is TOF_DISPODIST_MUL) then
    if Integer(Parms[1]) = 1 then TOF_DISPODIST_MUL(Latof).MULSiteDistant(True)
    else TOF_DISPODIST_MUL(Latof).MULSiteDistant(False);
end;
{$ENDIF}

procedure TOFDimensionsArticle(Parms: array of variant; nb: integer);
var F: TFMul;
  Latof: TOF;
begin
  F := TFMul(Integer(Parms[0]));
  if (F is TFMul) then Latof := TFMul(F).Latof else exit;
  if (Latof is TOF_DISPODIST_MUL) then TOF_DISPODIST_MUL(Latof).DimensionsArticle(string(Parms[1]), Integer(Parms[2]));
end;

procedure TOFNouvelleListe(Parms: array of variant; nb: integer);
{$IFNDEF EAGLCLIENT}
var F: TFMul;
  Latof: TOF;
  {$ENDIF}
begin
  {$IFNDEF EAGLCLIENT}
  F := TFMul(Integer(Parms[0]));
  if (F is TFMul) then Latof := TFMul(F).Latof else exit;
  if (Latof is TOF_DISPODIST_MUL) then TOF_DISPODIST_MUL(Latof).ChangeListe := True;
  {$ENDIF}
end;

initialization
  RegisterClasses([TOF_DISPODIST_MUL]);
  {$IFNDEF EAGLCLIENT}
  RegisterAglProc('ConnexionSiteDistant', True, 1, TOFConnexionSiteDistant);
  RegisterAglProc('ConnexionRNIS', True, 0, TOFConnexionRNIS);
  RegisterAglProc('AfficheParamSiteDistant', True, 0, TOFAfficheParamSiteDistant);
  RegisterAglProc('MULDistant', True, 1, TOFMULDistant);
  {$ENDIF}
  RegisterAglProc('DimensionsArticle', True, 2, TOFDimensionsArticle);
  RegisterAglProc('NouvelleListe', True, 0, TOFNouvelleListe);
end.
