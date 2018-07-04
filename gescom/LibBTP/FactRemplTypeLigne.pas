unit FactRemplTypeLigne;

interface
uses variants, sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,
     HCtrls,SaisUtil,HEnt1,UTOB,Menus,
     BTPUtil,paramsoc,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     CBPPath;

const MAXITEMS = 1;
type
  	TRemplTypeL = class (TObject)
    private
      fCreatedPop : boolean;
      fXX : TForm;
      fMaxItems : integer;
      fusable : boolean;
      POPGS : TPopupMenu;
    	MesMenuItem: array[0..MAXITEMS] of TMenuItem;
    	MesSMenuItem: array[0..2] of TMenuItem;
      procedure DefiniMenuPop;
      procedure ChangeTypeLTOComment (Sender : Tobject);
      procedure ChangeTypeLTOArt (Sender : Tobject);
      procedure ChangeTypeLTOSousTot (Sender : Tobject);
    public
      constructor create (FF : TForm);
      destructor destroy; override;
    	procedure Activate;
      procedure SetLigne (Arow : integer);
    end;

implementation
uses facture, Grids, FactTOB,FactComm,FactOuvrage,EntGC,FactUtil,FactArticle,UtilArticle,
  DB;


{ TRemplTypeL }


procedure TRemplTypeL.ChangeTypeLTOComment(Sender: Tobject);
var Arow,Acol : integer;
		XX : TFfacture;
    resp : integer;
    TOBL : TOB;
    Msg : string;
    ArtSav,LibSav : string;
begin
  XX := fXX as TFfacture;
	Arow := XX.GS.Row;
  Acol := XX.GS.Col;
  if XX.Action <> taConsult then
  begin
    XX.BouttonInVisible;
    TOBL := GetTOBLigne(XX.LaPieceCourante, ARow);
    if IsSousDetail(TOBL) then
    begin
      PgiError ('Impossible d''effectuer cette opération sur ce type de ligne');
      exit;
    end;
    ArtSav := TOBL.GetString('GL_REFARTSAISIE');
    libSav := TOBL.GetString ('GL_LIBELLE');
    if (TOBL.getValue('GL_ARTICLE')<>'') then
    begin
			Msg := 'ATTENTION : Cette ligne est valorisé et intervient dans le calcul de la pièce#13#10'+
      			 ' Etes-vous sur de définir cette ligne comme commentaire ?';
      resp := PGIAsk(msg);
      if resp <> Mryes then exit;
      XX.deduitLaLigne (TOBL);
      if IsOuvrage (TOBL) then
      begin
        XX.SuppressionDetailOuvrage (TOBL);
      end;
    end;
    XX.InitialiseLaligne(ARow);
    ZeroLigneMontant(TOBL);
    ReinitLigneFac (TOBL);
    TOBL.PutValue('GL_RECALCULER','X');
    XX.LaPieceCourante.PutValue('GP_RECALCULER','X');
    TOBL.SetString('GL_ARTICLE','');
    TOBL.SetString('GL_CODEARTICLE','');
    TOBL.SetString('GL_REFARTSAISIE',Artsav);
    TOBL.SetString('GL_LIBELLE',LibSav);
    XX.AfficheLaLigne(Arow);
    XX.GS.InvalidateRow(Arow);
    XX.GoToLigne(Arow,Acol);
    XX.CalculeLaSaisie(-1,-1,true,true,Arow,Arow);
    XX.BouttonVisible(Arow);
    XX.AnnuleSelection;
  end;
end;

procedure TRemplTypeL.ChangeTypeLTOSousTot(Sender: Tobject);
var Arow,Acol : integer;
		XX : TFfacture;
    resp : integer;
    TOBL : TOB;
    Msg : string;
    ArtSav,LibSav : string;
begin
  XX := fXX as TFfacture;
	Arow := XX.GS.Row;
  Acol := XX.GS.Col;
  if XX.Action <> taConsult then
  begin
    XX.BouttonInVisible;
    TOBL := GetTOBLigne(XX.LaPieceCourante, ARow);
    if IsSousDetail(TOBL) then
    begin
      PgiError ('Impossible d''effectuer cette opération sur ce type de ligne');
      exit;
    end;
    ArtSav := TOBL.GetString('GL_REFARTSAISIE');
    libSav := TOBL.GetString ('GL_LIBELLE');
    if (TOBL.getValue('GL_ARTICLE')<>'') then
    begin
			Msg := 'ATTENTION : Cette ligne est valorisé et intervient dans le calcul de la pièce#13#10'+
      			 ' Etes-vous sur de définir cette ligne comme commentaire ?';
      resp := PGIAsk(msg);
      if resp <> Mryes then exit;
      XX.deduitLaLigne (TOBL);
      if IsOuvrage (TOBL) then
      begin
        XX.SuppressionDetailOuvrage (TOBL);
      end;
    end;
    XX.InitialiseLaligne(ARow);
    ZeroLigneMontant(TOBL);
    ReinitLigneFac (TOBL);
    TOBL.SetString('GL_TYPELIGNE','TOT');
    TOBL.PutValue('GL_RECALCULER','X');
    XX.LaPieceCourante.PutValue('GP_RECALCULER','X');
    TOBL.SetString('GL_REFARTSAISIE',Artsav);
    TOBL.SetString('GL_LIBELLE',LibSav);
    XX.AfficheLaLigne(Arow);
    XX.GS.InvalidateRow(Arow);
    XX.GoToLigne(Arow,Acol);
    XX.CalculeLaSaisie(-1,-1,true,true,Arow,Arow);
    XX.BouttonVisible(Arow);
    XX.AnnuleSelection;
  end;
end;

procedure TRemplTypeL.Activate;
var indice : integer;
begin
  if (TFfacture(FXX).LaPieceCourante.GetValue('GP_NATUREPIECEG') <> VH_GC.AFNatProposition) then exit;
  for Indice := 0 to fMaxItems -1 do
  begin
  	if MesMenuItem [Indice] <> nil then
    begin
    	MesMenuItem[Indice].Visible := true;
    end;
  end;

end;

constructor TRemplTypeL.create(FF: TForm);
var ThePop : Tcomponent;
begin
  fXX := FF;
  fusable := false;
  fusable := true;
  ThePop := FF.Findcomponent  ('POPBTP');
  if ThePop = nil then
  BEGIN
    // pas de menu BTP trouve ..on le cree
    POPGS := TPopupMenu.Create(FF);
    POPGS.Name := 'POPBTP';
    fCreatedPop := true;
  END else
  BEGIN
    fCreatedPop := false;
    POPGS := TPopupMenu(thePop);
  END;
  DefiniMenuPop;
end;

procedure TRemplTypeL.DefiniMenuPop;
var Indice : integer;
begin
  fMaxItems := 0;
  if not fcreatedPop then
  begin
    MesMenuItem[fMaxItems] := TmenuItem.Create (fXX);
    with MesMenuItem[fMaxItems] do
    begin
      Caption := '-';
      Visible := false;
    end;
    inc (fMaxItems);
  end;
  // ---
  MesMenuItem[fMaxItems] := TmenuItem.Create (fXX);
  with MesMenuItem[fMaxItems] do
  begin
    Name := 'BCHANGETYPEL';
    Caption := TraduireMemoire ('Définir comme');
    Visible := false;
  end;

  MesSMenuItem [0] := TMenuItem.Create(fXX);
  MesSMenuItem [0].Name := 'BCHANTYPELCOMMENT';
  MesSMenuItem [0].Caption := TraduireMemoire ('Commentaire');
  MesSMenuItem [0].OnClick := ChangeTypeLTOComment;
  //
  MesSMenuItem [1] := TMenuItem.Create(fXX);
  MesSMenuItem [1].Name := 'BCHANTYPELCHIFFRAGE';
  MesSMenuItem [1].Caption := TraduireMemoire ('Ligne de chiffrage');
  MesSMenuItem [1].OnClick := ChangeTypeLTOArt;
  //
  MesSMenuItem [2] := TMenuItem.Create(fXX);
  MesSMenuItem [2].Name := 'BCHANTYPELSTOT';
  MesSMenuItem [2].Caption := TraduireMemoire ('Sous-Total');
  MesSMenuItem [2].OnClick := ChangeTypeLTOSousTot;

	MesMenuItem[fMaxItems].add (MesSMenuItem);
  inc (fMaxItems);

  for Indice := 0 to fMaxItems -1 do
  begin
  	if MesMenuItem [Indice] <> nil then
    begin
    	POPGS.Items.Add (MesMenuItem[Indice]);
    end;
  end;

end;

destructor TRemplTypeL.destroy;
var indice : integer;
begin
  for Indice := 0 to fMaxItems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
	if fcreatedPop then POPGS.free;
  inherited;
end;

procedure TRemplTypeL.SetLigne(Arow: integer);
var Indice : integer;
		TOBL : TOB;
    fvisi : boolean;
begin
  if Arow = 0 then exit;
  fVisi := false;
  TOBL := GetTOBLigne(TFFacture(fXX).LaPieceCourante, ARow);
  if IsArticle(TOBL) or (TOBL.getString('GL_TYPELIGNE')='TOT') or (TOBL.getString('GL_TYPELIGNE')='COM')  then fVisi := true;
  for Indice := 0 to fMaxItems -1 do
  begin
  	if MesMenuItem [Indice] <> nil then
    begin
    	MesMenuItem[Indice].Visible := fvisi;
    end;
  end;

end;

procedure TRemplTypeL.ChangeTypeLTOArt(Sender: Tobject);
var Arow,Acol : integer;
		XX : TFfacture;
    resp : integer;
    TOBL,TOBART : TOB;
    Msg : string;
    ArtSav,LibSav,Article,CodeArticle : string;
    QQ : Tquery;
begin
  ArtSav := '';
  XX := fXX as TFfacture;
	Arow := XX.GS.Row;
  Acol := XX.GS.Col;
  if XX.Action <> taConsult then
  begin
    XX.BouttonInVisible;
    TOBL := GetTOBLigne(XX.LaPieceCourante, ARow);
    if IsSousDetail(TOBL) or (IsArticle(TOBl)) or (IsOuvrage(TOBL)) then
    begin
      PgiError ('Impossible d''effectuer cette opération sur ce type de ligne');
      exit;
    end;

    if ISFromExcel(TOBL) then
    begin
      CodeArticle := GetParamSocSecur ('SO_BTARTICLEDIV','');
      if  CodeArticle = '' then
      begin
        PgiInfo ('Veuillez paramètrer l''article divers des appels d''offres');
        exit;
      end;
      QQ := OpenSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+CodeArticle+'"',True,1,'',True);

      Article := QQ.findField('GA_ARTICLE').AsString;
      ferme (QQ);
    end;
    ArtSav := TOBL.GetString ('GL_REFARTSAISIE');
    libSav := TOBL.GetString ('GL_LIBELLE');
    Msg := 'ATTENTION : Vous allez intégrer une ligne de chiffrage dans le document#13#10'+
           ' Valider cette opération ?';
    resp := PGIAsk(msg);
    if resp <> Mryes then exit;
    if ISFromExcel(TOBL) then
    begin
      TOBART := XX.TheTOBArticles.findFirst(['GA_ARTICLE'],[Article],false);
      if TOBART = nil then
      begin
        TOBArt := CreerTOBArt(XX.TheTOBArticles);
        QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                       'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                       'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+Article+'"',true,-1, '', True);
        TOBArt.SelectDB('',QQ);
        InitChampsSupArticle (TOBART);
        ferme (QQ);
      end;
      TOBL.SetString('GL_ARTICLE',Article);
      TOBL.SetString('GL_TYPELIGNE','ART');
      TOBL.SetString('GL_TYPEARTICLE',TOBArt.GetString('GA_TYPEARTICLE'));
      TOBL.SetString('GL_CODEARTICLE',TOBArt.GetString('GA_CODEARTICLE'));
      TOBL.SetString('GL_REFARTBARRE', TOBArt.GetString('GA_CODEBARRE'));
      TOBL.SetString('GL_PRIXPOURQTE', TOBArt.GetString('GA_PRIXPOURQTE'));
      TOBL.SetString('GL_TYPEREF', 'ART');
//      TOBL.SetDouble('GL_QTEFACT',1.0);
//      TOBL.SetDouble('GL_QTERESTE',1.0);
//      TOBL.SetDouble('GL_QTESTOCK',1.0);
//      TOBL.SetDouble('GL_QTERELIQUAT',1.0);
//      TOBL.SetDouble('GL_QTERELIQUAT',0);
     	TOBL.SetString('GL_RECALCULER','X');
      ArticleVersLigne (XX.LaPieceCourante ,TOBART,XX.TheTOBConds,TOBL,XX.TheTOBTiers);
      TraiteLesOuvrages(nil,XX.LaPieceCourante, XX.TheTOBArticles, XX.TheTOBOuvrage,nil,nil, Arow, False, XX.DEV, true);
    end else
    begin
    	XX.RechercheArticle (Arow);
    end;
    TOBL.PutValue('GL_RECALCULER','X');
    XX.LaPieceCourante.PutValue('GP_RECALCULER','X');
    if ArtSav <> '' then
    begin
      TOBL.SetString('GL_REFARTSAISIE',ArtSav);
      TOBL.SetString('GL_LIBELLE',LibSav);
    end;
    XX.AfficheLaLigne(Arow);
    XX.GS.InvalidateRow(Arow);
    XX.GoToLigne(Arow,Acol);
    XX.CalculeLaSaisie(-1,-1,true,true,Arow,Arow);
    XX.BouttonVisible(Arow);
    XX.AnnuleSelection;
  end;
end;

end.
