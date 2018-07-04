{***********UNITE*************************************************
Auteur  ...... : FV1 (VAUTRAIN)
Créé le ...... : 04/04/2013
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRECALCULPMAP ()
Mots clefs ... : TOF;RECALCULPMAP
*****************************************************************}
Unit UTOFBTRecalculPMAP ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Hpanel,
     HTB97,
     HRichOle,
     SAISUTIL,
     vierge,
     Grids,
     Graphics,
     Types,
     AffaireUtil,
     windows,
     Fe_main,
     UTOF,
     ///
     AglInit;

Type
  TOF_BTRECALCULPMAP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  Private
    fNomTable : string;
     //
    FF                : String;
    //
   	TOBGrille         : TOB;
    TobDepot          : TOB;
    TobArticle        : TOB;
    TobEmplacement    : TOB;
    TobDispo          : TOB;
    TobRecalcul       : TOB;
    //
    fColNamesGS       : string;
    Falignement       : string;
    Ftitre            : string;
    fLargeur          : string;
    //
    CodEmplacement    : String;
    //
    BLancer           : TToolbarButton97;
    BRevalorise       : TToolbarButton97;
    //
    GS                : THgrid;
    //
    CODEDEPOT         : THValComboBox;
    CODEARTICLE       : THEdit;
    DATERECALCUL      : THEdit;
    DateInvent        : THEdit;
    //
    LIBARTICLE        : THLabel;
    LIBDEPOT          : THLabel;
    LIBEMPLACEMENT    : THLabel;
    //
    DPA               : THNumEDit;
    PMAP              : THNumEDit;
    PMAPInvent        : THNumEdit;
    STOCKPHY          : THNumEDit;
    STOCKDISPO        : THNumEDit;
    STOCKNET          : THNumEDit;
    STOCKINVENT       : THNumEDit;
    //
    PMAPArticle       : Double;
    DPAArticle        : Double;
    PMAPDispo         : Double;
    DPADispo          : Double;
    //
    PMAPCalcUS        : Double;
    QTeSTockUS        : double;
    //
    procedure AfficheLaGrille(First: boolean=False);
    //
    procedure ArticleOnEnter(Sender: Tobject);
    procedure ArticleOnElipsisClick(Sender: Tobject);
    procedure ArticleOnExit(Sender: Tobject);
    //
    procedure BLancerOnClick(Sender: Tobject);
    Procedure BRevaloriseOnClick(Sender : TObject);
    //
    Procedure ChargeInfoArticle;
    procedure ChargeInfoDepot;
    Procedure ChargeInfoDispo;
    procedure ChargeInfoEmplacement;
    procedure ChargementInfoEntete;
    procedure ChargeTobGrille;
    //
    procedure ControleChamp(Champ, Valeur: String);
    procedure CreateTOB;
    //
    procedure DefinieGrid;
    procedure DessineGrille;
    procedure DestroyTOB;
    procedure DepotOnExit(Sender: Tobject);
    //
    procedure GestionArticle;
    procedure GestionDepot;
    procedure GetObjects;
    //
    procedure LectureInfoArticle;
    procedure LectureInfoEmplacement;
    //
    procedure RazZoneEcran;
    //
    procedure SetScreenEvents;
    function GetArticleFromCodeArticle(code: string): string;
    procedure LanceRevalorisation(TOBChoix: TOB);
    function GestionSiOk: boolean;
    function ISArticleOk: boolean;
    function ISDepotOk: boolean;
    procedure SetGrilleVide;
    procedure GSDblClick(Sender: TOBject);
    //
  end ;

Implementation
Uses MSGUtil,
     UtilPMAPCalcul,
     DateUtils,
     PiecesRecalculs,
     Facture;

procedure TOF_BTRECALCULPMAP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCULPMAP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCULPMAP.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCULPMAP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCULPMAP.OnArgument (S : String ) ;
var i       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
begin
  Inherited ;
  fNomTable := ConstitueNomTemp;
  //Chargement des zones ecran dans des zones programme
  GetObjects;

  CreateTOB;

  Critere := S;
               
  //Initialisation des zones écran.
  RazZoneEcran;

  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

  FF:='# ### ##';
  if V_PGI.OkDecV>0 then
  begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecV-1 do
    begin
       FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;
  SetControlVisible('HDATE',false);
  SetControlVisible('DATERECALCUL',false);
  //Gestion des évènement des zones écran
  SetScreenEvents;

  // définition de la grille
  DefinieGrid;

  //dessin de la grille
  DessineGrille;

  LIBDEPOT.Caption    := '';
  LIBARTICLE.Caption  := '';

  If (CodeDepot.Value <> '') and (CodeArticle.Text <> '') then ChargementInfoEntete;
  //
  BLancer.Enabled := false;
  BRevalorise.Enabled := false;
  //
end ;

Procedure TOF_BTRECALCULPMAP.ChargementInfoEntete;
begin

  //chargement des informations du Dépot (Entete)
  GestionDepot;

  GestionArticle;

end;


procedure TOF_BTRECALCULPMAP.OnClose ;
begin
  Inherited ;
  DropTableExist(fNomTable);
  DestroyTOB;
end ;

procedure TOF_BTRECALCULPMAP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCULPMAP.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCULPMAP.CreateTOB;
begin

  TOBDepot        := TOb.Create('DEPOT', nil, -1);
  TobArticle      := Tob.Create('ARTICLE',Nil, -1);
  TobEmplacement  := TOB.Create('EMPLACEMENT', nil, -1);
  TobDispo        := TOB.Create('DISPO', nil, -1);
  TobGrille       := Tob.create('GRILLE', nil,-1);

end;

procedure TOF_BTRECALCULPMAP.GetObjects;
begin
  //
  GS              := THgrid(GetControl('GRILLEPMAP'));
  //
  BLANCER         := TToolbarButton97(GetControl('BLANCER'));
  BRevalorise     := TToolBarButton97(GetCONTROL('BREVALORISATION'));
  //
  CODEDEPOT       := THValComboBox (GetControl('DEPOT'));
  //
  CODEARTICLE     := THEdit (GetControl('ARTICLE'));
  DATERECALCUL    := THEdit (GetControl('DATERECALCUL'));
  DATEINVENT      := THEdit (GetControl('DATEINV'));
  //
  DPA             := THNumEdit (GetControl('DPA'));
  PMAP            := THNumEdit (GetControl('PMAP'));
  PMAPINVENT      := THNumEdit (GetControl('PMAPINV'));
  STOCKPHY        := THNumEdit (GetControl('STOCKPHY'));
  STOCKDISPO      := THNumEdit (GetControl('STOCKDISPO'));
  STOCKNET        := THNumEdit (GetControl('STOCKNET'));
  STOCKINVENT     := THNumEdit (GetControl('STOCKINVENT'));
  //
  LIBARTICLE      := THLabel(GetControl('LIBARTICLE'));
  LIBDEPOT        := THLabel(GetControl('LIBDEPOT'));
  LIBEMPLACEMENT  := THLabel(GetControl('LIBEMPLACEMENT'));
  //
end;

procedure TOF_BTRECALCULPMAP.SetScreenEvents;
begin
  //
  BLANCER.OnClick             := BLancerOnClick;
  BRevalorise.OnClick         := BRevaloriseOnClick;
  //
  CODEDEPOT.OnExit            := DepotOnExit;
  //
  CodeArticle.OnElipsisClick  := ArticleOnElipsisClick;
  CodeArticle.OnExit          := ArticleOnExit;
  CodeArticle.OnEnter         := ArticleOnEnter;

end;

Procedure TOF_BTRECALCULPMAP.ArticleOnElipsisClick(Sender : Tobject);
Var stchamps,sWhere  : String;
begin
  CodeArticle.text := THEdit(getControl('ARTICLE')).text;
	sWhere := 'AND ((GA_TYPEARTICLE="MAR") OR (GA_TYPEARTICLE="ARP")) AND GA_TENUESTOCK="X"';
  if CodeArticle.Text <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArticle.text, 1, 18)) + ';XX_WHERE=' + sWhere
		 												else StChamps := 'XX_WHERE=' + sWhere;

	CodeArticle.text  := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps+';RECHERCHEARTICLE;STATUTART=UNI,DIM');
  ArticleOnExit(Self);
end;

Procedure TOF_BTRECALCULPMAP.DepotOnExit(Sender : Tobject);
begin
  If CODEDEPOT.Value <> '' then GestionDepot;
	GestionSiOk;
end;

Procedure TOF_BTRECALCULPMAP.GestionDepot;
begin

  //Lesture et chargement des informations du Dépot (Entete)
  tobdepot := LectureInfoDepot(CodeDepot.Value);
  ChargeInfoDepot;

  if CODEARTICLE.text <> '' then
  begin
    ChargeInfoDispo;

    //Lecture des informations de l'Emplacement (Entete)
    LectureInfoEmplacement;
    ChargeInfoEmplacement;
    //
  end;

end;

procedure TOf_BTRECALCULPMAP.GestionArticle;
begin
  if TOBarticle = nil then exit;
  //Lecture des informations de l'article (Entete)
  LectureInfoArticle;
  ChargeInfoArticle;

  //Lecture des informations du Dispo Article (Entete)
  LectureInfoDispo(CodeArticle.Text,CodeDepot.Value,TOBdispo);
  ChargeInfoDispo;

  //Lecture des informations de l'Emplacement (Entete)
  LectureInfoEmplacement;
  ChargeInfoEmplacement;

end;

function TOF_BTRECALCULPMAP.GetArticleFromCodeArticle (code : string) : string;
var QQ : TQuery;
begin
  result := '';
	QQ := OpenSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+ Code+'"',True,1,'',true);
  if not QQ.Eof then
  begin
    result := QQ.fields[0].AsString;
  end;
  ferme (QQ);
end;

function TOF_BTRECALCULPMAP.GestionSiOk : boolean;
begin
  if (CodeArticle.Text = '') or (CODEDEPOT.Value = '') then exit;
  if (IsArticleOK) and (IsDepotOk) then BLancer.Enabled := true;
end;


function TOF_BTRECALCULPMAP.ISArticleOk : boolean;
var CodeArt : string;
begin
  result := false;
  if not ExisteSql('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_ARTICLE="'+CodeArticle.text+'" AND GA_TYPEARTICLE <> "GEN"') then
  begin
    CodeArt := GetArticleFromCodeArticle (CodeArticle.Text);
    if CodeArt = '' then
    begin
      AfficheErreur(Ecran.name,'5', Ecran.Caption);
      Exit;
    end else
    begin
      CODEARTICLE.Text := CodeArt;
    end;
  end;
  GestionArticle;
  result := true;
end;

function TOF_BTRECALCULPMAP.ISDepotOk : boolean;
begin
  result :=  (CodeDepot.Value <> '');
end;

Procedure TOF_BTRECALCULPMAP.ArticleOnExit(Sender : Tobject);
begin
  if TOBArticle = nil then exit;
  if CodeDepot.Value = '' then
  begin
    CodeDepot.SetFocus;
    exit;
  end;

  //chargement des informations de l'article (Entete)
  if CodeArticle.Text <> '' then
  begin
    GestionSiOk;
  end;

end;

Procedure TOF_BTRECALCULPMAP.ArticleOnEnter(Sender : Tobject);
begin
	TOBARTICLE.InitValeurs;
end;

Procedure TOF_BTRECALCULPMAP.BLancerOnClick(Sender : Tobject);
var Date1  : TDateTime;
    Date2  : TdateTime;
begin
  GS.OnDblClick := GSDblClick; 
	if (CodeArticle.Text = '') or (CODEDEPOT.value='') then
  begin
    PGIError('Vous devez renseigner le dépot et le code de l''article..');
    exit;
  end;

  Date1  := StrToDate(DateInvent.Text);
  Date2  := StrToDate(DateRecalcul.text);

  ConstitueTableTemp(fNomTable,CodeArticle.Text,CODEDEPOT.Value,Date1,Date2);
  CalculsTableTempo(fNomTable,CodeDepot.Value, CodeArticle.Text,PMAPCalcUS,QTeSTockUS);

  ChargeTobGrille;
  AfficheLaGrille(True);
  BRevalorise.Enabled := (TOBGrille.detail.count > 0);
end;

Procedure TOF_BTRECALCULPMAP.ChargeTobGrille;
Var STSQL : String;
    QQ    : TQuery;
begin
  TobGrille.ClearDetail;

  StSql := 'SELECT * FROM '+fNomTable+' ORDER BY ONEDEPOT,ONEARTICLE,DATEPIECE,DATECREATION';
  QQ := OpenSQL(StSQL, False);
  If Not QQ.Eof then
  begin
    TobGrille.LoadDetailDB('CALCPMAP','','',QQ,False);
  end;
  if TOBGrille.detail.count = 0 then SetGrilleVide;
  ferme (QQ);

end;

Procedure TOF_BTRECALCULPMAP.AfficheLaGrille (First : boolean=false);
Var Indice : Integer;
begin
  GS.VidePile(false);
  if TOBGrille.detail.count = 0 then
		GS.RowCount :=  2
  else
		GS.RowCount := TOBGrille.detail.count + 1;

  GS.DoubleBuffered := true;
  GS.BeginUpdate;
	DessineGrille;

  TRY
    GS.SynEnabled := false;
    for Indice := 0 to TOBGrille.detail.count -1 do
    begin
      TOBGrille.Detail[Indice].PutLigneGrid (GS,Indice+1,false,false,fColNamesGS);
    end;
  FINALLY
    GS.SynEnabled := true;
    GS.EndUpdate;
  END;

  {if First then }TFVierge(ecran).HMTrad.ResizeGridColumns(GS);

end;

procedure TOF_BTRECALCULPMAP.DefinieGrid;
begin

  // Définition de la liste de saisie pour la grille Détail
  fColNamesGS := 'SEL;DATECREATION;NATUREPIECEG;NUMERO;LIBTIERS;DATEPIECE;QTEMVT;UNITEMVT;QTESTOCK;PU;PMAP;MTSTK';
  Falignement := 'C.0  ---;G.0  XX-;G.0 $---;C.0  ---;G.0  ---;C.0  -X-;D.3  -X-;G.0  ---;D.3  -X-;D.3  -X-;D.3  -X-;D.3  -X-';
  Ftitre      := '*;Date création;Nature;Numéro;Tiers;Date Mvt;Qté Mvt;Unité Mvt;Qté stock(US);PU Mvt;PMAP(US);Valo. Stock';
  fLargeur    := '2;16;25;20;70;16;10;10;16;20;16;20';

end;

procedure TOF_BTRECALCULPMAP.DessineGrille;
var st              : String;
    lestitres       : String;
    lesalignements  : String;
    alignement      : String;
    Nam             : String;
    leslargeurs     : String;
    lalargeur       : String;
    letitre         : String;
    lelement        : string;
    //
    Obli            : Boolean;
    OkLib           : Boolean;
    OkVisu          : Boolean;
    OkNulle         : Boolean;
    OkCumul         : Boolean;
    Sep             : Boolean;
    Okimg           : boolean;
    //
    dec             : Integer;
    NbCols          : integer;
    indice          : Integer;
begin
  //
  //Calcul du nombre de Colonnes du Tableau en fonction des noms
  st := fColNamesGS;

  NbCols := 0;

  repeat
    lelement := READTOKENST (st);
    if lelement <> '' then
    begin
      inc(NbCols);
    end;
  until lelement = '';
  //
  GS.ColCount     := Nbcols;
  //
  st              := fColNamesGS ;
  lesalignements  := Falignement;
  lestitres       := Ftitre;
  leslargeurs     := fLargeur;

  //Mise en forme des colonnes
  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom

    //
    alignement  := ReadTokenSt(lesalignements);
    lalargeur   := readtokenst(leslargeurs);
    letitre     := readtokenst(lestitres);

    //
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;

    //
    GS.cells[Indice,0]     := leTitre;
    GS.ColNames [Indice]   := Nam;

    //
    //Gestion des Zones Saisissables
    //

    //Alignement des cellules
    if copy(Alignement,1,1)='G'       then //Cadré à Gauche
      GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D'  then  //Cadré à Droite
      GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C'  then
      GS.ColAligns[indice] := taCenter; //Cadré au centre

    //Colonne visible ou non
    if OkVisu then
  		GS.ColWidths[indice] := strtoint(lalargeur)*GS.Canvas.TextWidth('W')
    else
    	GS.ColWidths[indice] := -1;

    //Affichage d'une image ou du texte
    okImg := (copy(Alignement,8,1)='X');
    if (OkLib) or (okImg) then
    begin
    	GS.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	GS.ColDrawingModes[Indice]:= 'IMAGE';
      end;
    end;

    //Gestion des décimales dans les zones de grille
    if (Dec<>0) or (Sep) then
    begin
    	if OkNulle then
        GS.ColFormats[indice] := FF+';; ;'
      else
      	GS.ColFormats[indice] := FF;
    end;
  end ;

end;

procedure TOF_BTRECALCULPMAP.ChargeInfoDepot;
begin

  CodeDepot.Value  := TobDepot.GetString('GDE_DEPOT');
  LIBDEPOT.Caption := TobDepot.GetString('GDE_LIBELLE');

end;


//lecture de la table Article pour affichage des éléments nécessaires
procedure TOF_BTRECALCULPMAP.LectureInfoArticle;
var StSQl : String;
    QQ    : TQuery;
begin
  if TOBArticle = nil then exit;
  StSQL := 'SELECT * FROM ARTICLE WHERE GA_ARTICLE="' + CodeArticle.text + '" AND GA_TYPEARTICLE <> "GEN"';
  QQ := OpenSQL(StSQL, False);
  If Not QQ.Eof then
  begin
    TobArticle.SelectDB('',QQ,False);
  end
  Else
  begin
    AfficheErreur(Ecran.name,'5', Ecran.Caption);
    CodeArticle.Text  := '';
    CodeArticle.SetFocus;
    Ferme(QQ);
    Exit;
  end;

  ferme (QQ);

end;

procedure TOF_BTRECALCULPMAP.ChargeInfoArticle;
begin


  CodeArticle.Text    := TobArticle.GetString('GA_ARTICLE');
  LIBARTICLE.Caption  := TobArticle.GetString('GA_LIBELLE');
  //
  DPAArticle          := TobArticle.GetDouble('GA_DPA');
  PMAPArticle         := TobArticle.GetDouble('GA_PMAP');
  THLabel(GetControl('UNITESTOCK')).Caption := RechDom('GCQUALUNITPIECE',TobArticle.GetString('GA_QUALIFUNITESTO'),false)
end;

procedure TOF_BTRECALCULPMAP.ChargeInfoDispo;
begin

  If TobDispo = nil then exit;

  DateRecalcul.text := DateToStr(now);
  DateInvent.text   := DateToStr(TobDispo.GetDateTime('GQ_DATEINV'));
  PMAPInvent.text   := FloatToSTr(TobDispo.GetDouble('GQ_PRIXINV'));
  //
  CodEmplacement    := TobDispo.GetString('GQ_EMPLACEMENT');
  //
  PMAPDispo := TobDispo.GetDouble('GQ_DPA');
  DPADispo  := TobDispo.GetDouble('GQ_PMAP');
  //
  DPA.Text  := FloatToSTr(TobDispo.GetDouble('GQ_DPA'));
  PMAP.Text := FloatToSTr(TobDispo.GetDouble('GQ_PMAP'));
  //
  StockPhy.Text     := FloatToSTr(Tobdispo.GetDouble('GQ_PHYSIQUE'));
  StockDispo.Text   := FloatToSTr(Tobdispo.GetDouble('GQ_QTEDISPO'));
  StockNet.Text     := FloatToSTr(Tobdispo.GetDouble('GQ_QTENET'));
  StockInvent.Text  := FloatToSTr(Tobdispo.GetDouble('GQ_STOCKINV'));

end;

Procedure TOF_BTRECALCULPMAP.RazZoneEcran;
begin
  //
  CODEDEPOT.text   := '';
  CodeArticle.Text := '';
  //
  LibDepot.caption        := '';
  LibArticle.caption      := '';
  LibEmplacement.caption  := '';
  //
  CodEmplacement    := '';
  //
  DateRecalcul.Text := DateToStr(Now);
  Dateinvent.Text   := DateToStr(iDate1900);
  PMAPInvent.Text   := FloatToStr(0);
  //
  DPA.Text  := FloatToStr(0.00);
  PMAP.Text := FloatToStr(0.00);
  //
  StockPhy.Text     := FloatToStr(0);
  StockDispo.Text   := FloatToStr(0);
  StockNet.Text     := FloatToStr(0);
  StockInvent.Text  := FloatToStr(0);
  //
  PMAPdispo         := 0;
  PMAPArticle       := 0;
  DPAArticle        := 0;
  PMAPArticle       := 0;
  //
end;

//lecture de la table DISPO (GD_) pour affichage des éléments de Stocks
procedure TOF_BTRECALCULPMAP.LectureInfoEmplacement;
var StSQl : String;
    QQ    : TQuery;
begin
  if TOBemplacement = nil then exit;
  StSQL := 'Select * FROM EMPLACEMENT WHERE GEM_DEPOT="' + CodeDepot.Value + '" AND GEM_EMPLACEMENT="' + CodEmplacement + '"';
  QQ := OpenSQL(StSQL, False);
  If Not QQ.Eof then
  begin
    TobEmplacement.SelectDB('',QQ,False);
  end;

  ferme (QQ);

end;

procedure TOF_BTRECALCULPMAP.ChargeInfoEmplacement;
begin

  If TobEmplacement = nil then exit;

  LIBEMPLACEMENT.Caption := TOBEmplacement.GetString('GEM_LIBELLE');

end;


procedure TOF_BTRECALCULPMAP.ControleChamp(Champ, Valeur: String);
begin

  if champ ='DEPOT'   then CODEDEPOT.Text := Valeur;

  if champ ='ARTICLE' then CodeARTICLE.Text := Valeur;

end;

procedure TOF_BTRECALCULPMAP.DestroyTOB;
begin

  FreeAndNil(TOBDepot);
  FreeAndNil(TOBArticle);
  FreeAndNil(TOBEmplacement);
  FreeAndNil(TobDispo);
  FreeAndNil(TobRecalcul);
  FreeAndNil(TOBGrille);

end;

procedure TOF_BTRECALCULPMAP.LanceRevalorisation(TOBChoix : TOB);
Var Q1 : TQuery;
    Ind: Integer;
    StSql : string;
    TOBMAj,TOBPieces : TOB;
begin
  TOBMAj := TOB.Create('UNE LIV',nil,-1);
  TOBPieces := TOB.Create ('LES LIGNES',nil,-1);
  //lecture de la table tempo pour recalcul PMAP
  StSql := 'SELECT * FROM '+fNomTable+' ORDER BY ONEDEPOT,ONEARTICLE,DATEPIECE,DATECREATION';
  Q1 := OpenSQL(StSQL, False);
  Q1.First;

  TRY
    while Not Q1.EOF do
    begin
      TobMaj.SelectDB('LPMAP',Q1);
      //Mise à jour des livraison
      IF (TobMaj.GetString('NATUREPIECEG')='LBT') OR (TobMaj.GetString('NATUREPIECEG')='BLC') then
      begin
        if TOBCHOIX.geTString('LES LIV')='X' then MiseAjourPMAPLiv(TOBMAJ,TOBPieces);
        //mise à jour des consommations
        if TOBCHOIX.geTString('LES CONSOS')='X' then MAjPMAPConso(TOBMAJ);
      end;
      Q1.Next;
    end;
  FINALLY
    Ferme (Q1);
    MiseAjourDispoSuiteInv(CODEDEPOT.Value ,CODEARTICLE.Text,QTeSTockUS,PMAPCalcUS);
    //recalcul des pieces de livraisons modifiees
    if Assigned(TOBPieces) then
    begin
      For ind := 0 to TOBPieces.detail.count-1 do
      begin
        if TraitementRecalculPiece(TOBPieces.Detail[ind],false,false)<>TrrOk  then V_PGI.ioerror := OeUnknown;
      end;
    end;
    TobMAJ.free;
    TOBPieces.free;
  END;

end;

procedure TOF_BTRECALCULPMAP.BRevaloriseOnClick(Sender: TObject);
var TOBChoix : TOB;
begin
	TOBChoix := TOB.Create ('LES CHOIX',nil,-1);
  TRY
    TOBChoix.AddChampSupValeur('OKGO','-');
    TOBChoix.AddChampSupValeur('LES LIV','-');
    TOBChoix.AddChampSupValeur('LES CONSOS','-');
    TheTOB := TOBChoix;
    AGLLANCEFICHE('BTP', 'BTREVALPMAP','','','ARTICLE='+CodeArticle.text+';DEPOT='+CODEDEPOT.Value);
    if TheTOB <> nil then TheTOB := nil;
    if (TOBChoix.GetString('OKGO')='X') then
    begin
    	LanceRevalorisation(TOBChoix);
    end;
  FINALLY
    TOBChoix.free;
  END;
  //remise à zéro de l'écran
  RazZoneEcran;
  //
  TobGrille.ClearDetail;
	SetGrilleVide;
  //
  AfficheLaGrille;
  //
  CodeDepot.SetFocus;
  BLancer.enabled := false;
  BRevalorise.Enabled := false;
  //
end;

procedure TOF_BTRECALCULPMAP.SetGrilleVide;
begin
  TOBGrille.AddChampSupValeur('TYPELIGNE','');
  TOBGrille.AddChampSupValeur('NATUREPIECEG','');
  TOBGrille.AddChampSupValeur('SOUCHE','');
  TOBGrille.AddChampSupValeur('NUMERO',0);
  TOBGrille.AddChampSupValeur('NOLIGNE',0);
  TOBGrille.AddChampSupValeur('NUMORDRE',0);
  TOBGrille.AddChampSupValeur('ONEDEPOT',0);
  TOBGrille.AddChampSupValeur('LIBDEPOT',0);
  TOBGrille.AddChampSupValeur('ONEARTICLE',0);
  TOBGrille.AddChampSupValeur('LIBARTICLE',0);
  TOBGrille.AddChampSupValeur('ONETIERS',0);
  TOBGrille.AddChampSupValeur('LIBTIERS',0);
  TOBGrille.AddChampSupValeur('DATECREATION',0);
  TOBGrille.AddChampSupValeur('DATEPIECE',0);
  TOBGrille.AddChampSupValeur('PMAP',0);
  TOBGrille.AddChampSupValeur('PMAPUV',0);
  TOBGrille.AddChampSupValeur('GPP_ESTAVOIR','');
  TOBGrille.AddChampSupValeur('UNITEACH','');
  TOBGrille.AddChampSupValeur('COEFUAUS',0);
  TOBGrille.AddChampSupValeur('UNITEVTE','');
  TOBGrille.AddChampSupValeur('COEFUSUV',0);
  TOBGrille.AddChampSupValeur('PUSTK',0);
  TOBGrille.AddChampSupValeur('QTESTK',0);
  TOBGrille.AddChampSupValeur('UNITESTO','');
  TOBGrille.AddChampSupValeur('UNITEMVT','');
  TOBGrille.AddChampSupValeur('MTSTK',0);
end;

procedure TOF_BTRECALCULPMAP.GSDblClick(Sender: TOBject);
var NaturePiece,Souche : string;
		Numero,Indice : Integer;
    cledoc : String;
    TOBL : TOB;
begin
  if TOBgrille.detail.count < GS.row  then exit;
  TOBL := TOBGrille.detail[GS.row-1]; if TOBL = nil then exit;
  if TOBL.GetValue('NATUREPIECEG')='' then exit;

	NaturePiece:=TOBL.GetValue('NATUREPIECEG');
	Souche:=TOBL.GetValue('SOUCHE');
	Numero:=TOBL.GetValue('NUMERO');
	Indice:=0;
  cledoc := Naturepiece+';;'+Souche+';'+InttoStr(Numero)+';'+InttoStr(Indice)+';';
  AppelPiece([cledoc,'ACTION=CONSULTATION'],2);
end;

Initialization
  registerclasses ( [ TOF_BTRECALCULPMAP ] ) ; 
end.

