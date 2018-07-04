{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/10/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISIECONSOMOB ()
Mots clefs ... : TOF;BTSAISIECONSOMOB
*****************************************************************}
Unit BTSAISIECONSOMOB_TOF;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Forms,
     Graphics,
     ComCtrls,
     DbCtrls,
{$IFDEF EAGLCLIENT}
      eMul,
      Maineagl,
      EFiche,
      UtileAGL,
      HQry,
{$ELSE}
      db,
      {$IFNDEF DBXPRESS}
      dbTables,
      {$ELSE}
      uDbxDataSet,
      {$ENDIF}
      DBGrids,
      mul,
      FE_Main,
      Fiche,
{$ENDIF}
      sysutils,
      HCtrls,
      HEnt1,
      HMsgBox,
      Hpanel,
      HTB97,
      HRichOle,
      vierge,
      Types,
      windows,
      UTOF,
      HDB,
      ParamSoc,
      ParamDBG,
      UTOM,
      MsgUtil,
      HSysMenu,
      AglInit,
      utob;

Type
  TOF_BTSAISIECONSOMOB = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  Private
    //
    CodeAppel   : String;
    NumMouv     : String;
    Recherche   : String;
    CritRech    : String;
    CodeRecup   : String;
    Article     : String;
    TypeAction  : String;
    Zoneget     : string;
    ZONEORDRE   : string;
    //
    Hierarchie  : Boolean;
    //
    TobLigConso : Tob;
    TobArticle  : Tob;
    TOBECHANGE  : TOB;
    //
    LFam1       : THLabel;
    LFam2       : THLabel;
    LFam3       : THLabel;
    LLibFamArt1 : THLabel;
    LLibFamArt2 : THLabel;
    LLibFamArt3 : THLabel;
    //
    DateMouv    : THEdit;
    FamART1     : THEdit;
    FamART2     : THEdit;
    FamART3     : THEdit;
    CodeArt     : THEdit;
    //
    LibArt      : THEdit;
    QteArt      : THNumEdit;
    //
    PBouton     : TToolWindow97;
    //
    Dock971     : TDock97;
    //
    BQuitter    : TToolBarButton97;
    BValider    : TToolBarButton97;
    BDelete     : TToolBarButton97;
    BRecherche  : TToolBarButton97;
    //

    procedure BDelete_Click(Sender: TObject);
    procedure BRecherche_Click(Sender: TObject);
    procedure BValider_Click(Sender: TObject);    
    procedure ChargeEcran(ZoneRech: String);
    procedure ChargeLigneEcran;
    procedure ChargementTOBArticle;
    procedure CodeArt_OnEnter(Sender: Tobject);
    procedure ControleChamp(Champ, Valeur: String);
    Procedure CreateTOB;
    procedure DesTroyTob;
    procedure FamArt1_OnEnter(Sender: Tobject);
    procedure FamArt2_OnEnter(Sender: Tobject);
    procedure FamArt3_OnEnter(Sender: Tobject);
    procedure GetObjects;
    procedure InitEcran;
    procedure LibArt_OnEnter(Sender: Tobject);
    procedure SetScreenEvents;
    procedure Zone_Exit(Sender: TObject);
    procedure ActiveZoneEcran;
    function LectureTOBArticle: boolean;
    function FindTOBArticle(CodeArticle: string): boolean;
    //
  end ;


Implementation
uses UtilSaisieConso;

procedure TOF_BTSAISIECONSOMOB.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIECONSOMOB.OnDelete ;
begin
  Inherited ;
  TOBEchange.PutValue('VALIDE', 'S');
end ;

procedure TOF_BTSAISIECONSOMOB.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIECONSOMOB.OnLoad ;
begin
  Inherited ;

  ChargementTOBArticle;

end ;

procedure TOF_BTSAISIECONSOMOB.OnArgument (S : String ) ;
var X       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;
  //
  TOBECHANGE := LaTOB;
  Hierarchie := GetParamSoc('SO_GCFAMHIERARCHIQUE');

  CreateTOB;
  
  //traitement Arguments
	Critere :=(Trim(ReadTokenSt(S)));

  while (Critere <> '') do
  begin
    if Critere <> '' then
    begin
      X := pos (':', Critere) ;
      if x = 0 then
        X := pos ('=', Critere) ;
      if x <> 0 then
      begin
        Champ := copy (Critere, 1, X - 1) ;
        Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        ControleChamp(champ, valeur);
      end
    end;
    Critere := (Trim(ReadTokenSt(S)));
  end;

  //Chargement des objets écran dans l'uses
  GetObjects;

  //remise à zéro des zone écran
  InitEcran;

  //chargement de la ligne conso si en modif
  ChargeLigneEcran;

  ActiveZoneEcran;

  DateMouv.SetFocus;

  //Définition des Evènement des zones écrans
  SetScreenEvents;


end ;

Procedure TOF_BTSAISIECONSOMOB.ChargeLigneEcran;
begin

  if NumMouv <> '' then
  begin
    //TobLigConso := TOBEchange.FindFirst(['BCO_NUMMOUV'],[NumMouv], False);
    If TobEchange.GetString('CODEART') = '' then
    begin
      TypeAction := 'C';
      BDelete.Visible := False;
      BValider.Visible := False;
      if ASsigned(DateMouv)     Then DateMouv.Text        := DateToStr(V_PGI.DateEntree);
    end
    Else
    begin
      TypeAction := 'M';
      BDelete.Visible := True;
      BValider.Visible := True;
      if ASsigned(DateMouv)     Then DateMouv.Text        := DateToStr(TobEchange.GetDateTime('DATEMOUV'));
      //
      Article := TobEchange.GetString('ARTICLE');
      if Assigned(LibArt)       Then LibArt.Text          := TobEchange.GetString('LIBART');
      //
      LectureTobArticle;
      //
      //ChargementTOBArticle;
      //
      if Assigned(QteArt)       Then QteArt.Value          := TobEchange.GetDouble('QTEART');
    end;
  end
  else
  begin
    TypeAction := 'C';
    BDelete.Visible := False;
    BValider.Visible := False;
    if ASsigned(DateMouv)       Then DateMouv.Text        := DateToSTr(V_PGI.DateEntree);
  end;

end;

procedure TOF_BTSAISIECONSOMOB.OnClose ;
begin

  DestroyTob;

  Inherited ;
end ;

procedure TOF_BTSAISIECONSOMOB.DesTroyTob () ;
begin

  FreeAndNil(TobLigConso);

end ;


procedure TOF_BTSAISIECONSOMOB.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISIECONSOMOB.OnCancel () ;
begin
  Inherited ;

  FreeAndNil(TobArticle);

end ;

procedure TOF_BTSAISIECONSOMOB.ControleChamp(Champ, Valeur: String);
begin
  if Champ = 'APPEL'    Then CodeAppel := Valeur;
  //
  if Champ = 'NUMMOUV'  Then NumMouv := Valeur;
end;

Procedure TOF_BTSAISIECONSOMOB.GetObjects;
begin

  if assigned(GetControl('PBouton'))        then PBouton      := TToolWindow97(GetControl('PBOUTON'));
  if Assigned(GetControl('Dock971'))        then Dock971      := TDock97(Getcontrol('Dock971'));
  //
  if Assigned(GetControl('BDELETE'))        then BDelete      := TToolBarButton97(GetControl('BDELETE_'));
  if Assigned(GetControl('BRECHERCHE'))     then BRecherche   := TToolBarButton97(GetControl('BRECHERCHE'));
  if Assigned(GetControl('BQUITTER'))       then BQuitter     := TToolBarButton97(GetControl('BQUITTER'));
  if Assigned(GetControl('BVALIDER'))       then BValider     := TToolBarButton97(GetControl('BVALIDER_'));
  //
  if Assigned(Getcontrol('LFAM1'))          then LFam1        := THLabel(GetControl('LFAM1'));
  if Assigned(Getcontrol('LFAM2'))          then LFam2        := THLabel(GetControl('LFAM2'));
  if Assigned(Getcontrol('LFAM3'))          then LFam3        := THLabel(GetControl('LFAM3'));
  //
  if Assigned(Getcontrol('LLIBFAMART1'))    then LLibFamArt1  := THLabel(GetControl('LLIBFAMART1'));
  if Assigned(Getcontrol('LLIBFAMART2'))    then LLibFamArt2  := THLabel(GetControl('LLIBFAMART2'));
  if Assigned(Getcontrol('LLIBFAMART3'))    then LLibFamArt3  := THLabel(GetControl('LLIBFAMART3'));
  //
  if Assigned(Getcontrol('DATEMOUV'))       then DateMouv     := THEdit(GetControl('DATEMOUV'));
  if Assigned(Getcontrol('FAMART1'))        then FamArt1      := THEdit(GetControl('FAMART1'));
  if Assigned(Getcontrol('FAMART2'))        then FamArt2      := THEdit(GetControl('FAMART2'));
  if Assigned(Getcontrol('FAMART3'))        then FamArt3      := THEdit(GetControl('FAMART3'));
  if Assigned(Getcontrol('CODEART'))        then CodeArt      := THEdit(GetControl('CODEART'));
  if Assigned(Getcontrol('LIBART'))         then LibArt       := THEdit(GetControl('LIBART'));
  if Assigned(Getcontrol('QTEART'))            then QteArt       := THNumEdit(GetControl('QTEART'));

end;

procedure TOF_BTSAISIECONSOMOB.Zone_Exit(Sender :TObject);
begin
  BDelete.Visible := false;
  BValider.Visible := false;

  if ThEdit(Sender).Name = 'CODEART' then
  begin
		if not FindTOBArticle (CodeArt.Text) then BEGIN CodeARt.SetFocus ;Exit; END;
    LectureTobArticle;
    ChargementTOBArticle;
  end;
  if (CodeArt.text <> '') Or (LibArt.Text<> '') Or (QteArt.Value <> 0) then
  begin
    BDelete.Visible := True;
    BValider.Visible := True;
  end;
end;

procedure TOF_BTSAISIECONSOMOB.SetScreenEvents;
begin

  if Assigned(BDelete)        then BDelete.OnClick        := BDelete_Click;
  if Assigned(BRecherche)     then BRecherche.OnClick     := BRecherche_Click;
  If Assigned(BValider)       then BValider.OnClick       := BValider_Click;
  //
  If assigned(CodeArt)        then CodeArt.OnExit         := Zone_Exit;
  If assigned(LibArt)         then LibArt.OnExit          := Zone_Exit;
  If assigned(QteArt)         then QteArt.OnExit          := Zone_Exit;
  //
  If assigned(Famart1)        then FamArt1.OnEnter        := FamArt1_OnEnter;
  If assigned(Famart2)        then FamArt2.OnEnter        := FamArt2_OnEnter;
  If assigned(Famart3)        then FamArt3.OnEnter        := FamArt3_OnEnter;
  If assigned(CodeArt)        then CodeArt.OnEnter        := CodeArt_OnEnter;
  If assigned(LibArt)         then LibArt.OnEnter         := LIBArt_OnEnter;
  //

end;

Procedure TOF_BTSAISIECONSOMOB.FamArt1_OnEnter(Sender : Tobject);
Begin

  Recherche:= '';
  CritRech := '';
  CodeRecup := '';
  ZONEORDRE := 'CC_CODE';

  Zoneget := 'FAMART1';
  Recherche := 'BTRECHFAMNIV1MOB';
  CritRech  := '';
  CodeRecup := 'CC_CODE';

end;

Procedure TOF_BTSAISIECONSOMOB.FamArt2_OnEnter(Sender : Tobject);
Begin

  Recherche:= '';
  CritRech := '';
  CodeRecup := '';
  ZONEORDRE := 'CC_CODE';

  Zoneget := 'FAMART2';
  Recherche := 'BTRECHFAMNIV2MOB';
  if Hierarchie then CritRech  := 'CC_LIBRE LIKE "' + FamART1.Text + '%"';
  CodeRecup := 'CC_CODE';

end;

Procedure TOF_BTSAISIECONSOMOB.FamArt3_OnEnter(Sender : Tobject);
Begin

  Recherche:= '';
  CritRech := '';
  CodeRecup := '';
  ZONEORDRE := 'CC_CODE';

  Zoneget := 'FAMART3';
  Recherche := 'BTRECHFAMNIV3MOB';
  if Hierarchie then CritRech  := 'CC_LIBRE LIKE "' + FamART1.text+FamART2.Text + '%"';
  CodeRecup := 'CC_CODE';

end;

Procedure TOF_BTSAISIECONSOMOB.CodeArt_OnEnter(Sender : Tobject);
Begin

  Recherche:= '';
  CritRech := '';
  CodeRecup := '';

  Zoneget := 'CODEART';
  Recherche := 'BTRECHARTMOBILE';
  CritRech  := 'GA_TYPEARTICLE IN ("MAR", "PRE", "ARP") ';
  ZONEORDRE := 'GA_CODEARTICLE';

  if FamArt1.Text <> '' then CritRech := CritRech + 'AND GA_FAMILLENIV1="' + FamArt1.Text + '" ';
  if FamArt2.Text <> '' then CritRech := CritRech + 'AND GA_FAMILLENIV2="' + FamArt2.Text + '" ';
  if FamArt3.Text <> '' then CritRech := CritRech + 'AND GA_FAMILLENIV3="' + FamArt3.Text + '" ';

  CodeRecup := 'GA_ARTICLE';
end;

Procedure TOF_BTSAISIECONSOMOB.LibArt_OnEnter(Sender : Tobject);
Begin
  Recherche:= '';
  CritRech := '';
  CodeRecup := '';
  Zoneget := 'LIBART';
  ZONEORDRE := 'GA_LIBELLE';

  Recherche := 'BTRECHARTMOBILE';
  CritRech  := 'GA_TYPEARTICLE IN ("MAR", "PRE","ARP") ';
  if FamArt1.Text <> '' then CritRech := CritRech + 'AND GA_FAMILLENIV1="' + FamArt1.Text + '" ';
  if FamArt2.Text <> '' then CritRech := CritRech + 'AND GA_FAMILLENIV2="' + FamArt2.Text + '" ';
  if FamArt3.Text <> '' then CritRech := CritRech + 'AND GA_FAMILLENIV3="' + FamArt3.Text + '" ';
  CodeRecup := 'GA_ARTICLE';
end;

Procedure TOF_BTSAISIECONSOMOB.InitEcran;
begin
  //
  if Assigned(LFam1)        then LFam1.caption := '';
  if Assigned(LFam2)        then LFam2.caption := '';
  if Assigned(LFam3)        then LFam3.caption := '';
  //
  if Assigned(LFam1)        Then LFam1.caption := RechDom('GCLIBFAMILLE', 'LF1', True);
  if Assigned(LFam2)        Then LFam2.caption := RechDom('GCLIBFAMILLE', 'LF2', True);
  if Assigned(LFam3)        Then LFam3.caption := RechDom('GCLIBFAMILLE', 'LF3', True);
  //
  if Assigned(FamArt1)      then FamArt1.Text  := '';
  if Assigned(FamArt2)      then FamArt2.Text  := '';
  if Assigned(FamArt3)      then FamArt3.Text  := '';
  //
  if Assigned(CodeArt)      then CodeArt.Text  := '';
  if Assigned(LibArt)       then LibArt.Text   := '';
  if Assigned(QteArt)       then QteArt.Value   := 0;
  //
end;

procedure TOF_BTSAISIECONSOMOB.BDelete_Click(Sender: TObject);
begin

  TOBECHANGE.PutValue('VALIDE', 'S');

end;

procedure TOF_BTSAISIECONSOMOB.BRecherche_Click(Sender: TObject);
var TOBRecherche  : TOB;
begin

  if (recherche='') or (Coderecup='') then exit;

	CodeArt.OnExit := nil;

  TOBRecherche := TOB.Create('RECHERCHE', Nil, -1);
  TobRecherche.AddChampSupValeur('CODE', '');
  TheTob := Tobrecherche;

  AGLLanceFiche('BTP', 'BTRECHMOB', '', '','LISTEPARAM=' + Recherche + '; '+
  									   'CODERECUP=' + CodeRecup + '; '+
                       'CRITRECH=' + CritRech+';'+
                       'SAISINIT='+GetControltext(Zoneget)+';'+
                       'ORDERBY='+ZONEORDRE);

  ChargeEcran(Tobrecherche.GetString('CODE'));
	CodeArt.OnExit := Zone_Exit;

end;

Procedure TOF_BTSAISIECONSOMOB.ChargeEcran(ZoneRech : String);
begin

  if ZoneRech = '' then
  begin
    exit;
  end;

  if Recherche = 'BTRECHFAMNIV1MOB'     then
  begin
    if Assigned(FamArt1)        Then
    begin
      FamArt1.text := ZoneRech;
      if Assigned(LLibFamArt1)  Then LLibFamArt1.Caption  := RechDom('GCFAMILLENIV1',FamArt1.Text, True);
    end;
  end
  else if Recherche = 'BTRECHFAMNIV2MOB'  then
  begin
    if Assigned(FamArt1)        Then
    begin
      FamArt2.text := ZoneRech;
      if Assigned(LLibFamArt2)  Then LLibFamArt2.Caption  := RechDom('GCFAMILLENIV2',FamArt2.Text, True);
    end;
  end
  else if Recherche = 'BTRECHFAMNIV3MOB'  then
  begin
    if Assigned(FamArt3)        Then
    begin
      FamArt3.text := ZoneRech;
      if Assigned(LLibFamArt3)  Then LLibFamArt3.Caption  := RechDom('GCFAMILLENIV3',FamArt3.Text, true);
    end;
  end
  else if Recherche = 'BTRECHARTMOBILE'   then
  begin
    if Assigned(CodeArt)        Then
    begin
      Article := ZoneRech;
      LectureTobArticle;
      ChargementTOBArticle;
      if Article <> '' then
      begin
        BDelete.Visible := true;
        BValider.Visible := true;
      end;
    end;
    QTeArt.SetFocus;
  end
  else
    PGIError('La Recherche n''a rien donnée', 'Recherche');
end;

Procedure TOF_BTSAISIECONSOMOB.ChargementTOBArticle;
Begin

  FamArt1.text          := TOBArticle.GetString('GA_FAMILLENIV1');
  if FamArt1.Text <> '' then LLibFamArt1.Caption   := RechDom('GCFAMILLENIV1',FamArt1.Text, True);
  //
  FamArt2.text          := TOBArticle.GetString('GA_FAMILLENIV2');
  if FamArt2.Text <> '' then LLibFamArt2.Caption   := RechDom('GCFAMILLENIV2',FamArt2.Text, True);
  //
  FamArt3.text          := TOBArticle.GetString('GA_FAMILLENIV3');
  if FamArt3.Text <> '' then LLibFamArt3.Caption   := RechDom('GCFAMILLENIV3',FamArt3.Text, true);

  if TobArticle.GetString('GA_ARTICLE') <> '' then
  begin
    CodeArt.Text  := TobArticle.GetString('GA_CODEARTICLE');
    Article       := TobArticle.GetString('GA_ARTICLE');
  	LibArt.Text   := TobArticle.GetString('GA_LIBELLE');
  end;
end;

function TOF_BTSAISIECONSOMOB.FindTOBArticle (CodeArticle : string) : boolean;
var QQ : TQuery;
		StSql : string;
begin
  result := false;
  StSQL := 'SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="' + CodeArticle + '" ';
  QQ := OpenSql (StSql,true,1,'',True);
  if not QQ.eof then BEGIN Article := QQ.findField('GA_ARTICLE').AsString; Result := true; END;
  ferme (QQ);
end;

function TOF_BTSAISIECONSOMOB.LectureTOBArticle : boolean;
Var StSQL     : String;
    QArticle  : TQuery;
begin
	Result := True;
  StSQL := 'SELECT * FROM ARTICLE WHERE GA_ARTICLE="' + Article + '" ';

  if FamArt1.Text <> '' then StSQL := StSql + '  AND GA_FAMILLENIV1="' + FamArt1.Text + '" ';
  if FamArt2.Text <> '' then StSQL := StSql + '  AND GA_FAMILLENIV2="' + FamArt2.Text + '" ';
  if FamArt3.Text <> '' then StSQL := StSql + '  AND GA_FAMILLENIV3="' + FamArt3.Text + '" ';

  QArticle := OpenSQL(StSQL, True,1,'',true);

  if QArticle.Eof then
  begin
    //AfficheErreur('',4);
    Result := false;
    Ferme(QArticle);
    exit;
  end;

  TobArticle.SelectDB('', QArticle);

  Ferme(QArticle);

end;

procedure TOF_BTSAISIECONSOMOB.CreateTOB;
begin

  TobArticle := TOB.Create('ARTICLE', Nil, -1);

end;

procedure TOF_BTSAISIECONSOMOB.BValider_Click(Sender: TObject);
begin

  //
  NextControl(TfVierge(Ecran),true);
  // Controle avant validation
  if QteArt.value = 0 then
  begin
    Ecran.modalresult := 0;
    QTeArt.SetFocus;
    Exit;
  end;

  //Chargement des zones TOBEchange pour récupération TOBArticle et écran de saisie
  TOBECHANGE.PutValue('DATEMOUV', DateMouv.text);

  TOBECHANGE.PutValue('FamArt1', FamART1.text);
  TOBECHANGE.PutValue('FamArt2', FamART2.text);
  TOBECHANGE.PutValue('FamArt3', FamART3.text);
  TOBECHANGE.PutValue('CODEART', CodeArt.text);
  TOBECHANGE.PutValue('LIBART',  LibArt.text);
  TOBECHANGE.PutValue('QTEART',  QteArt.Value);

  TOBECHANGE.PutValue('VALIDE', TypeAction);

  TOBECHANGE.PutValue('TYPEARTICLE', TobArticle.GetString('GA_TYPEARTICLE'));

  TobArticle.ChangeParent(TOBECHANGE, -1);
end;

procedure TOF_BTSAISIECONSOMOB.ActiveZoneEcran;
begin

  if TypeAction = 'C' then
  begin
    FamArt1.enabled := True;
    FamArt2.enabled := True;
    FamArt3.enabled := True;
    CodeArt.enabled := True;
    LibArt.enabled  := True;
    BRecherche.Visible := True;
  end
  else
  begin
    FamArt1.enabled := False;
    FamArt2.enabled := False;
    FamArt3.enabled := False;
    CodeArt.enabled := False;
    BRecherche.Visible := False;
  end;

end;

Initialization
  registerclasses ( [ TOF_BTSAISIECONSOMOB ] ) ;
end.

