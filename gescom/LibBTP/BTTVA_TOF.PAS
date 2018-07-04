{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 24/02/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTTVA ()
Mots clefs ... : TOF;BTTVA
*****************************************************************}
Unit BTTVA_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
     Types,
     Vierge,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,fe_main,
{$else}
     eMul, MainEagl,
{$ENDIF}
		 grids,
     uTob,
		 HCtrls,
		 HTB97,
     forms,
     sysutils,
     ComCtrls,
     HEnt1,
     HMsgBox,
     EntGc,
     Ent1,
     UTOF,
     ParamSoc,
     Hcompte,
     Hpanel,
     UtilPgi,
     Messages,
     printDbg,
     Windows,uEntCommun
     ;

const
  //
  ListeChampsRTV = 'SEL;CC_CODE;CC_LIBELLE;';
  flargeurRTV = '2;6;30;';
  FalignementRTV = 'G.0  ---;G.0  ---;G.0  ---;';
  FTitreRTV = ' ; Code ; Libellé ;';
  NCRTV = '0;0;0;'; //definition des zones saisissable ou non
  //
  TexteMessage: array[1..20] of string  = (

{1}     'Vous devez renseigner un code',
{2}     'Vous devez renseigner un libellé',
{3}     'Le code que vous avez saisi existe déjà. Vous devez le modifier',
{4}     'L''enregistrement est inaccessible',
{5}     'Le compte à l''achat débit que vous avez saisi n''existe pas, est incompatible ou incomplet',
{6}     'Le compte à la vente débit que vous avez saisi n''existe pas, est incompatible ou incomplet.',
{7}     'Vous devez renseigner le taux à l''achat',
{8}     'Vous devez renseigner le taux à la vente',
{9}     'Vous devez renseigner le compte à l''achat',
{10}    'Vous devez renseigner le compte à la vente',
{11}    'L''enregistrement ne sera pas sauvegardé',
{12}    'Le code est référencé par un compte général',
{13}    'Le compte à l''achat encaissement que vous avez saisi n''existe pas, est incompatible ou incomplet',
{14}    'Le compte à la vente encaissement que vous avez saisi n''existe pas, est incompatible ou incomplet',
{15}    'Cet indice de code TVA est déjà attribué',
{16}    'Le compte de TVA de R.G. que vous avez saisi n''existe pas, est incompatible ou incomplet',
{17} 		'Voulez-vous enregistrer les modifications?',
{18} 		'Veuillez renseigner une désignation pour ce régime de TVA',
{19} 		'Etes-vous sur de vouloir supprimer ce régime de TVA ?',
{20} 		'Suppression Impossible.Le code est référencé par un compte général.'
    );

Type
  TOF_BTTVA = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	ModeCreat : boolean;
  	PHAUT : THpanel;
  	ErrorGs : boolean;
  	StCellCur : STRING;
    Cache: THCpteEdit;
  	TOBRTV : TOB; // Liste des regimes de TVA
    TOBTAXTVA : TOB; // Table des taux et comptes de TVA
    TOBR : TOB;
    GRTV,GS : THgrid;
    CREGIME,CTAUXACH,CTAUXVTE,CCPTEACH,CCPTEVTE,CENCAISACH,CENCAISVTE,CCPTEVTERG,CCPTECOLLEC : integer; // colonnes de saisie
    CC_CODE,CC_LIBELLE,CC_LIBRE : THEdit;
    VALLIBRE : THSpinEdit;
    //
    ListeChampsGS : string;
    flargeurGS : string;
    FalignementGS : string;
    FTitreGS : string;
    NCGS : string;
    //
    BTVAETAB,Binsert,Bdelete,BImprimer,Bvalider : TToolbarButton97;
  	procedure AlloueTobs;
    procedure libereTobs;
    procedure GetControls;
    procedure DefinieGrids;
    procedure ChargeGridRTV;
    procedure definiGridGS;
    procedure definiGridRTV;
    procedure InitGrille(Grid: Thgrid; FieldList, Ftitre, Flargeur,Falignement,FcolsEditable: string; OnlySelect, TwoColors: boolean);
    procedure TraiteChampAffaireGrid(Gr: THGrid; Nom: string;Col: integer);
		procedure PositionneGRTV (Ligne : integer);
    procedure SetEvents (etat : boolean);
    procedure SetEventsRTV (etat : boolean);
    procedure SetEventsGS (Etat : boolean);
    procedure AfficheGridRTV;
    procedure AfficheGridGS;
    procedure PositionneColsGS;
    procedure AfficheInfoRegime (TOBL : TOB);
    procedure BinsertClick (Sender : TObject);
    //
    procedure ChargeGS (TOBL : TOB);
    // gestion de la grille GRTV
    procedure GRTVPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas;AState: TGridDrawState);
    procedure GRTVRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GRTVRowExit(Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean);
    // gestion de la grille GS
    procedure GSEnter(Sender: TObject);
    procedure GSExit(Sender: TObject);
    procedure GSDblClick(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSElipsisClick(Sender: TObject);
    // --
    function ChercheCpte(var Cpte: string; Vide: Boolean): byte;
    function ControlDonnee (Acol,Arow : integer) : boolean;
    procedure SetDonneeLigne (Acol,Arow : integer);
    function FabReqTva(Cpt: string): string;
    procedure GotoLigne (row,col : integer);
    function ISDonneesOK(Arow: integer; var Acol: integer): boolean;
    function ControleGrilleGS : boolean;
    function EcritDonnes : boolean;
    function controleEntete: boolean;
    procedure MajTXCPTA;
    procedure CodeExit (Sender : Tobject);
    procedure LibelleExit (Sender : TObject);
    procedure VALLIBREExit (Sender : TObject);
    procedure BdeleteClick (Sender : Tobject);
    procedure BimprimerClick (Sender : Tobject);
    procedure BTVAETABClick (Sender : TObject);
    procedure AjouteLesNouveauxRegimes(TOBTAXTVA, TOBR: TOB);
  end ;

procedure ParamTVATPF(OkTva: boolean);

Implementation

uses TntGrids,TVAEtab;

procedure ParamTVATPF(OkTva: boolean);
begin
  if _Blocage(['nrCloture', 'nrBatch', 'nrSaisieModif', 'nrEnca', 'nrDeca'], True, 'nrBatch') then Exit;
  AGLLanceFiche('BTP','BTTVA','','','');
  _Bloqueur('nrBatch', False);
end;


procedure TOF_BTTVA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTTVA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTTVA.OnUpdate ;
begin
  Inherited ;
  GSExit(self);
  if not ErrorGs then
  begin
    if TOBTAXTVA.IsOneModifie(true) then
    begin
      if not EcritDonnes then
      BEGIN
        ecran.ModalResult := 0;
      	EXIT;
      END else
      begin
        if TOBRTV.detail[GRTV.row-1].getValue('NEW')='X' then
        begin
          TOBRTV.detail[GRTV.row-1].putValue('CC_TYPE',VH^.DefCatTVA);
          TOBRTV.detail[GRTV.row-1].putValue('CC_CODE',CC_CODE.text);
          TOBRTV.detail[GRTV.row-1].putValue('CC_LIBELLE',CC_LIBELLE.text);
          TOBRTV.detail[GRTV.row-1].putValue('CC_ABREGE',COPY(CC_LIBELLE.text,1,17));
          TOBRTV.detail[GRTV.row-1].putValue('CC_LIBRE',CC_LIBRE.text);
          TOBRTV.detail[GRTV.row-1].InsertDB(nil);
          TOBRTV.detail[GRTV.row-1].putValue('NEW','-');
          TOBRTV.detail[GRTV.row-1].SetAllModifie(false);
  				TOBRTV.detail[GRTV.row-1].PutLigneGrid(GRTV,GRTV.row,false,false,ListeChampsRTV);
        end;
      end;
    end;
  end;
	CHARGETVATPF;
  TFVierge(ecran).ModalResult := 0;
end ;

procedure TOF_BTTVA.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTTVA.OnArgument (S : String ) ;
var QQ : TQuery;
    NomChoixCod : String;
begin
  Inherited ;

  if getparamSocSecur('SO_BTVENTCOLLECTIF',false) then
  begin
    ListeChampsGS := 'SEL;TV_REGIME;TV_TAUXACH;TV_TAUXVTE;TV_CPTEACH;TV_CPTEVTE;TV_COLLECTIF;TV_ENCAISACH;TV_ENCAISVTE;TV_CPTVTERG;';
    flargeurGS := '2;5;7;7;12;12;12;12;12;12;';
    FalignementGS := 'G.0  ---;G.0O ---;D.2  -X-;D.2  -X-;C.0  ---;C.0  ---;C.0  ---;C.0  ---;C.0  ---;C.0  ---;';
    FTitreGS := ' ;Rég.;Tx Ach;Tx Vte;Cpte Achat;Cpte Vente;Collectif;Cpte Achat Encais.;Cpte Vente Encais.;Cpte Tva RG;';
    NCGS := '0;0;1;1;1;1;1;1;1;1;'; //definition des zones saisissable ou non
  end else
  begin
    ListeChampsGS := 'SEL;TV_REGIME;TV_TAUXACH;TV_TAUXVTE;TV_CPTEACH;TV_CPTEVTE;TV_ENCAISACH;TV_ENCAISVTE;TV_CPTVTERG;';
    flargeurGS := '2;5;7;7;12;12;12;12;12;';
    FalignementGS := 'G.0  ---;G.0O ---;D.2  -X-;D.2  -X-;C.0  ---;C.0  ---;C.0  ---;C.0  ---;C.0  ---;';
    FTitreGS := ' ;Rég.;Tx Ach;Tx Vte;Cpte Achat;Cpte Vente;Cpte Achat Encais.;Cpte Vente Encais.;Cpte Tva RG;';
    NCGS := '0;0;1;1;1;1;1;1;1;'; //definition des zones saisissable ou non
  end;

  ModeCreat := false;
	Cache := THCpteEdit.Create(ecran);
  Cache.Parent := ecran;
  Cache.visible := false;
  AlloueTobs;

  // Récupération nom de base principale en cas de partage
  NomChoixcod := 'CHOIXCOD';
  QQ := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="TTREGIMETVA"', True);
  if not QQ.EOF then
  begin
  	NomChoixCod := QQ.FindField('DS_NOMBASE').AsString+'.DBO.CHOIXCOD';
  end;
	Ferme (QQ);

  QQ := OpenSql('Select CC_CODE From '+NomChoixCod+' Where CC_TYPE="RTV" Order by CC_CODE', True,-1,'',true);
  TOBR.LoadDetailDb ('CHOIXCOD','','',QQ,false);
  ferme (QQ);

  GetControls;
  DefinieGrids;
  ChargeGridRTV;
  AfficheGridRTV;
  SetEventsRTV (true);
  PositionneGRTV (1);
end ;

procedure TOF_BTTVA.OnClose ;
begin
  Inherited ;
  //
  //
  Cache.free;
  libereTobs;
end ;

procedure TOF_BTTVA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTTVA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTTVA.ChargeGridRTV;
var QQ : Tquery;
		Indice : integer;
    NomChoixCod : String;
begin
  // Récupération nom de base principale en cas de partage
  NomChoixcod := 'CHOIXCOD';
  QQ := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="TTTVA"', True);
  if not QQ.EOF then
  begin
  	NomChoixCod := QQ.FindField('DS_NOMBASE').AsString+'.DBO.CHOIXCOD';
  end;
	Ferme (QQ);

	QQ := OpenSql('SELECT * FROM '+NomChoixCod+' WHERE CC_TYPE="TX1" ORDER BY CC_LIBRE', True,-1,'',true);
  TOBRTV.LoadDetailDB('CHOIXCOD','','',QQ,false);
  ferme (QQ);
  for Indice := 0 to TOBRTV.detail.count -1 do
  begin
  	TOBRTV.detail[Indice].AddChampSupValeur('SEL','');
  	TOBRTV.detail[Indice].AddChampSupValeur('NEW','-');
  end;
end;

procedure TOF_BTTVA.InitGrille (Grid : Thgrid; FieldList,Ftitre,Flargeur,Falignement,FcolsEditable : string; OnlySelect,TwoColors : boolean);
var Col,Dec, Larg : integer ;
    tal : TAlignment ;
    STitre,St,Ch,stA,FF,Typ,FcolEditable : string ;
    Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul : boolean ;
    Nom : string;
    lescolonnes : string;
    Indice : integer;
    ColName : string;
begin
  Col:=0; STitre:=''; ColName:='';

  With TFMul(Ecran) do
  BEGIN
    GS.Titres.Add(FieldList) ;
    While Ftitre<> '' do
    BEGIN
      StA:=ReadTokenSt(FAlignement);
      St:=ReadTokenSt(Ftitre);
      Ch:=ReadTokenSt(FieldList);
      Larg:=ReadTokenI(FLargeur);
      tal:=TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
      FcolEditable:=ReadTokenSt(FcolsEditable);
      if OkVisu then
      BEGIN
      	if Col >= Grid.colCount then Grid.colCount := Grid.ColCount + 1;
        Grid.Cells[Col,0]:=St ;
        Grid.ColAligns[Col]:=tal;
        Grid.ColWidths[Col]:=Larg*GS.Canvas.TextWidth('W') ;
        if (OkLib) or (ch='TV_COLLECTIF') then
        begin
          if ch='TV_COLLECTIF' then
          begin
            Grid.ColFormats[Col]:='CB=GCCPTGENERAUX|AND G_GENERAL LIKE "41%"|<<Défaut>>';
          end else
          begin
            Grid.ColFormats[Col]:='CB=' + Get_Join(Ch);
          end;
        end;
        Typ:=ChampToType(Ch) ;
        if (Typ='INTEGER') or (Typ='SMALLINT') or (Typ='DOUBLE') then Grid.ColFormats[Col]:=FF ;
        if Stitre = '' then Stitre := St else STitre:=STitre+';'+St;
        if Colname = '' then Colname:= ch else ColName:=ColName+';'+Ch;
        if FcolEditable = '0' then Grid.ColEditables [Col]:=false;
        inc (Col);
      END;
    END;
  END ;
  //
  lescolonnes := ColName;
  Nom := READTOKENST (LesColonnes);
  Indice := 0;
  repeat
    TraiteChampAffaireGrid (Grid,Nom,indice);
    Nom := READTOKENST (LesColonnes);
    inc(Indice);
  until nom='';
  if OnlySelect then
  begin
    Grid.Options:=Grid.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
    Grid.Options:=Grid.Options+[GoRowSelect] ;
  	Grid.MultiSelect := true;
  end;
  Grid.TwoColors := TwoColors;
end;

procedure TOF_BTTVA.definiGridRTV;
var FieldList,Ftitre,Flargeur,Falignement,FcolEditable : string;
begin
  FieldList := ListeChampsRTV;
  Ftitre := FTitreRTV;
  Flargeur := flargeurRTV;
  Falignement := FalignementRTV;
  FcolEditable := NCRTV;
  InitGrille (GRTV,FieldList,Ftitre,Flargeur,Falignement,FcolEditable,True,true);
end;

procedure TOF_BTTVA.definiGridGS;
var FieldList,Ftitre,Flargeur,Falignement,FcolEditable : string;
		FF : string;
    i : integer;
begin
  FieldList := ListeChampsGS;
  Ftitre := FTitreGS;
  Flargeur := flargeurGS;
  Falignement := FalignementGS;
  FcolEditable := NCGS;

  InitGrille (GS,FieldList,Ftitre,Flargeur,Falignement,FcolEditable,false,false);
	PositionneColsGS;
  //
  for i := 1 to VH^.Cpta[fbGene].Lg do FF := FF + 'a';
  FF :=  '>' + FF + ';0; ';
	GS.ColFormats[CCPTEACH]:=FF ;
	GS.ColFormats[CCPTEVTE]:=FF ;
	GS.ColFormats[CENCAISACH]:=FF ;
	GS.ColFormats[CENCAISVTE]:=FF ;
	GS.ColFormats[CCPTEVTERG]:=FF ;
  //
  FF := '#0.00';
	GS.ColFormats[CTAUXACH]:=FF ;
	GS.ColFormats[CTAUXVTE]:=FF ;
end;

procedure TOF_BTTVA.DefinieGrids;
begin
	definiGridRTV;
	definiGridGS;
end;

procedure TOF_BTTVA.GetControls;
begin
  GRTV := THgrid(GetControl('GRTV'));
  GS := THgrid(GetControl('GS')) ;
  CC_CODE := Thedit(getCOntrol('CC_CODE'));
  CC_LIBELLE := Thedit(getCOntrol('CC_LIBELLE'));
  CC_LIBRE := Thedit(getCOntrol('CC_LIBRE'));
  BTVAETAB := TToolbarButton97 (GetControl('BTVAETAB'));
  Binsert := TToolbarButton97 (GetControl('Binsert'));
  Bdelete := TToolbarButton97 (GetControl('Bdelete'));
  BImprimer := TToolbarButton97 (GetControl('BImprimer'));
  Bvalider := TToolbarButton97 (GetControl('Bvalider'));
  VALLIBRE := THSpinEdit(GetControl('VALLIBRE'));
  //
  BTVAETAB.onclick := BTVAETABClick;
  Binsert.onclick := BinsertClick;
  Bdelete.onClick := BdeleteClick;
  BImprimer.onclick := BimprimerClick;
  PHAUT := THpanel(GetCOntrol('PHAUT'));
end;

procedure TOF_BTTVA.AlloueTobs;
begin
  TOBRTV := TOB.create ('LES REGIMES TVA',nil,-1);
  TOBTAXTVA := TOB.create ('LES CPTE TVA',nil,-1);
  TOBR := TOB.create ('LES DIFF TVA',nil,-1);
end;

procedure TOF_BTTVA.libereTobs;
begin
  TOBRTV.free;
  TOBTAXTVA.free;
  TOBR.free;
end;

Procedure TOF_BTTVA.TraiteChampAffaireGrid ( Gr : THGrid; Nom : string; Col : integer);
Var  Lg,ind,NumPart : integer;
     NomChamp,Libelle : string;
     Visible : Boolean;
begin
  if nom = '' then Exit;
  lg := Length(Nom); ind:= Pos ('_',Nom);
  if (lg =0) or (ind = 0) then  exit;
  Nomchamp := Copy (Nom, ind+1,lg-ind);
  Visible:=False; //mcd 14/02/03
  If ((Nomchamp = 'AFFAIRE1') Or (Nomchamp = 'AFFAIRE2') Or (Nomchamp = 'AFFAIRE3')) Then
  Begin
    NumPart := StrToInt(Copy(NomChamp, 8,1));
    Case NumPart Of
      1 : Begin Visible:=True; Libelle:=VH_GC.CleAffaire.Co1Lib;  End;
      2 : Begin Visible:=VH_GC.CleAffaire.Co2Visible; Libelle:=VH_GC.CleAffaire.Co2Lib; End;
      3 : Begin Visible:=VH_GC.CleAffaire.Co3Visible; Libelle:=VH_GC.CleAffaire.Co3Lib; End;
    End;
    if not(Visible) then Gr.colwidths[Col]:=0;
    Gr.Cells[Col,0]:=Libelle;
  End;
  if ((Nomchamp='AVENANT') and Not(VH_GC.CleAffaire.GestionAvenant)) then Gr.colwidths[Col]:=0;
end;

procedure TOF_BTTVA.PositionneGRTV(Ligne: integer);
var cancel : boolean;
begin
	cancel := false;
	GRTV.Row := Ligne;
  GRTVRowEnter(self,Ligne,cancel,false);
end;

procedure TOF_BTTVA.GRTVRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
begin
	GRTV.InvalidateRow(Ou);
  SetEventsGS (false); // desactive la gestion d'evenement de la grille de saisie
  TOBL := TOBRTV.detail[Ou-1];
  AfficheInfoRegime (TOBL);
  ChargeGS (TOBL);
  GS.SetFocus;
  GS.row := 1;
//uniquement en line
//GS.Col := CTAUXVTE;
  GS.Col := CTAUXACH;
  GSEnter (self);
  SetEventsGS (true);
  GS.ShowEditor;
end;

procedure TOF_BTTVA.GRTVRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
	if ErrorGS then BEGIN cancel := true; exit; END;
//	SendMessage(GS.Handle, WM_KeyDown, VK_DOWN, 0);
  if not ControleGrilleGS then BEGIN cancel := true; EXIT; END;
  if TOBTAXTVA.IsOneModifie(true) then
  begin
  	if PgiAsk (TexteMessage[17])=Mryes then
    begin
  		if not EcritDonnes then BEGIN cancel := true; EXIT; END;
      if TOBRTV.detail[Ou-1].getValue('NEW')='X' then
      begin
        TOBRTV.detail[Ou-1].putValue('CC_TYPE',VH^.DefCatTVA);
        TOBRTV.detail[Ou-1].putValue('CC_CODE',CC_CODE.text);
        TOBRTV.detail[Ou-1].putValue('CC_LIBELLE',CC_LIBELLE.text);
        TOBRTV.detail[Ou-1].putValue('CC_ABREGE',COPY(CC_LIBELLE.text,1,17));
        TOBRTV.detail[Ou-1].putValue('CC_LIBRE',CC_LIBRE.text);
        TOBRTV.detail[Ou-1].InsertDB(nil);
        TOBRTV.detail[Ou-1].putValue('NEW','-');
        TOBRTV.detail[Ou-1].SetAllModifie(false);
        TOBRTV.detail[Ou-1].PutLigneGrid(GRTV,Ou,false,false,ListeChampsRTV);
      end;
    end else
    begin
    	if TOBRTV.detail[Ou-1].getValue('NEW')='X' then
      begin
        TOBRTV.detail[Ou-1].free;
        GRTV.CacheEdit;
        GRTV.SynEnabled := False;
        GRTV.DeleteRow(Ou);
        GRTV.RowCount := GRTV.RowCount -1;
        GRTV.MontreEdit;
        GRTV.SynEnabled := true;
      end;
    end;
  end;
  GRTV.InvalidateRow(Ou);
  SetEvents(false);
  PHAUT.enabled := false;
end;

procedure TOF_BTTVA.GRTVPostDrawCell(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var ARect: TRect;
		Triangle: array[0..2] of TPoint;
begin
  if (gdFixed in AState) and (ACol = 0) then
  begin
    Arect := GRTV.CellRect(Acol, Arow);
    Canvas.Brush.Color := GRTV.FixedColor;
    Canvas.FillRect(ARect);
    if (ARow = GRTV.row) then
    begin
      Canvas.Brush.Color := clBlack;
      Canvas.Pen.Color := clBlack;
      Triangle[1].X := ARect.Right - 2;
      Triangle[1].Y := ((ARect.Top + ARect.Bottom) div 2);
      Triangle[0].X := Triangle[1].X - 5;
      Triangle[0].Y := Triangle[1].Y - 5;
      Triangle[2].X := Triangle[1].X - 5;
      Triangle[2].Y := Triangle[1].Y + 5;
      if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle);
    end;
  end;
end;

procedure TOF_BTTVA.SetEventsRTV (etat : boolean);
begin
	if Etat then
  begin
    GRTV.OnRowEnter := GRTVRowEnter;
    GRTV.OnRowExit := GRTVRowExit;
    GRTV.PostDrawCell := GRTVPostDrawCell;
  end else
  begin
		GRTV.OnRowEnter := nil;
    GRTV.OnRowExit := nil;
    GRTV.PostDrawCell := nil;
  end;
end;

procedure TOF_BTTVA.AfficheGridRTV;
var Indice : integer;
begin
	GRTV.RowCount := TOBRTV.detail.count+1;
	for Indice := 0 to TOBRTV.detail.count -1 do
  begin
  	TOBRTV.detail[Indice].PutLigneGrid(GRTV,Indice+1,false,false,ListeChampsRTV);
  end;
  TFvierge(ecran).HMTrad.ResizeGridColumns(GRTV);
end;

procedure TOF_BTTVA.SetEventsGS(Etat: boolean);
begin
	if Etat then
  begin
  	GS.OnEnter := GSEnter;
  	GS.OnExit := GSExit;
  	GS.OnRowEnter := GSRowEnter;
  	GS.OnCellEnter := GSCellEnter;
  	GS.OnRowExit := GSRowExit;
  	GS.OnCellExit := GSCellExit;
  	GS.OnDblClick  := GSDblClick;
    GS.OnElipsisClick := GSElipsisClick;
  end else
  begin
  	GS.OnEnter := nil;
  	GS.OnExit := nil;
  	GS.OnRowEnter := nil;
  	GS.OnCellEnter := nil;
  	GS.OnRowExit := nil;
  	GS.OnCellExit := nil;
  	GS.OnDblClick  := nil;
    GS.OnElipsisClick := nil;
  end;
end;


procedure TOF_BTTVA.AjouteLesNouveauxRegimes (TOBTAXTVA,TOBR : TOB);
var Indice : integer;
		TOBN : TOB;
begin
  for Indice := 0 to TOBR.detail.count -1 do
  begin
    if TOBTAXTVA.FindFirst(['TV_REGIME'],[TOBR.detail[Indice].GetString ('CC_CODE')],True) = nil then
    begin
    	TOBN := TOB.Create ('TXCPTTVA',TOBTAXTVA,-1);
      TOBN.putValue('TV_REGIME',TOBR.detail[Indice].GetString('CC_CODE'));
      TOBN.putValue('TV_TVAOUTPF',VH^.DefCatTVA);
    	TOBN.PutValue('TV_CODETAUX',CC_CODE.text);
    end;
  end;
  TOBTAXTVA.Detail.Sort('TV_REGIME');
end;

procedure TOF_BTTVA.ChargeGS(TOBL: TOB);
var QQ : TQuery;
		Indice : integer;
    TOBN : TOB;
    Req, NomBase : String;
begin
	TOBTAXTVA.clearDetail;

  // Récupération nom de base principale en cas de partage
  NomBase := '';
  QQ := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="TTREGIMETVA"', True);
  if not QQ.EOF then
  begin
  	NomBase := QQ.FindField('DS_NOMBASE').AsString+'.DBO.';
  end;
	Ferme (QQ);

  Req := 'SELECT * FROM '+Nombase+'TXCPTTVA WHERE TV_TVAOUTPF="'+VH^.DefCatTVA+'" AND TV_CODETAUX="'+TOBL.getValue('CC_CODE')+'"';
  QQ := OpenSql (Req,true,-1,'',true);
  if not QQ.eof then
  begin
    TOBTAXTVA.LoadDetailDB('TXCPTTVA','','',QQ,false);
    AjouteLesNouveauxRegimes (TOBTAXTVA,TOBR);
    for Indice := 0 to TOBTAXTVA.detail.count -1 do
    begin
      TOBTAXTVA.detail[Indice].AddChampSupValeur ('SEL','');
    end;
  	TOBTAXTVA.SetAllModifie(FALSE);
  end else
  begin
    for Indice := 0 to TOBR.detail.count -1 do
    begin
    	TOBN := TOB.Create ('TXCPTTVA',TOBTAXTVA,-1);
      TOBN.putValue('TV_REGIME',TOBR.detail[Indice].getValue('CC_CODE'));
      TOBN.putValue('TV_TVAOUTPF',VH^.DefCatTVA);
    	TOBN.PutValue('TV_CODETAUX',CC_CODE.text);
    end;
  end;
  AfficheGridGS;
end;

procedure TOF_BTTVA.AfficheGridGS;
var Indice : integer;
begin
	GS.RowCount := TOBTAXTVA.detail.count+1;
	for Indice := 0 to TOBTAXTVA.detail.count -1 do
  begin
  	TOBTAXTVA.detail[Indice].PutLigneGrid(GS,Indice+1,false,false,ListeChampsGS);
  end;
  GS.Invalidate;
end;

procedure TOF_BTTVA.PositionneColsGS;
var Nb : integer;
		LeChamp : string;
    LaListe : string;
begin
	Nb := 0;
	laListe := ListeChampsGS;
  repeat
    LeChamp := READTOKENST (laListe);
    if LeChamp <> '' then
    begin
      if LeChamp = 'TV_REGIME' then  CREGIME := Nb else
      if LeChamp = 'TV_TAUXACH' then CTAUXACH := Nb else
      if LeChamp = 'TV_TAUXVTE' then CTAUXVTE := Nb else
      if LeChamp = 'TV_CPTEACH' then CCPTEACH := Nb else
      if LeChamp = 'TV_CPTEVTE' then CCPTEVTE := Nb else
      if LeChamp = 'TV_COLLECTIF' then CCPTECOLLEC := Nb else
      if LeChamp = 'TV_ENCAISACH' then CENCAISACH := Nb else
      if LeChamp = 'TV_ENCAISVTE' then CENCAISVTE := Nb else
      if LeChamp = 'TV_CPTVTERG' then CCPTEVTERG := Nb else
    end;
    inc (nb);
    until LeChamp = '';
    GS.ColCount := Nb-1;
    if (not GetParamSocSecur('SO_GCVENTILRG', False)) or
       (not GetParamSocSecur('SO_RETGARANTIE', False)) then
    begin
      // Supression dans la grid des champs associées
      GS.ColLengths [CCPTEVTERG] := -1;
      GS.ColWidths  [CCPTEVTERG] := 0;
    end;
//uniquement en line
{* 	GS.ColLengths [CCPTEACH] := -1; GS.ColWidths  [CCPTEACH] := 0;
  	GS.ColLengths [CTAUXACH] := -1; GS.ColWidths  [CTAUXACH] := 0;
  	GS.ColLengths [CCPTEACH] := -1; GS.ColWidths  [CCPTEACH] := 0;
  	GS.ColLengths [CENCAISACH] := -1; GS.ColWidths  [CENCAISACH] := 0;
*}
end;

procedure TOF_BTTVA.AfficheInfoRegime(TOBL: TOB);
var Valeur : integer;
begin
	if TOBL.getValue('CC_LIBRE') = '' then TOBL.PutValue('CC_LIBRE','0');
  if not IsNumeric(TOBL.getValue('CC_LIBRE')) then TOBL.PutValue('CC_LIBRE','0');
  Valeur := StrToInt(TRim(TOBL.getValue('CC_LIBRE')));
  VALLIBRE.Value := Valeur;
  CC_CODE.text := TOBL.getValue('CC_CODE');
  CC_LIBELLE.text := TOBL.getValue('CC_LIBELLE');
end;

procedure TOF_BTTVA.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
	if (GS.Col = CCPTEACH) or
  	 (GS.Col = CCPTEVTE) or
     (GS.Col = CENCAISACH) or
     (GS.Col = CENCAISVTE) or
     (GS.Col = CCPTEVTERG) then GS.ElipsisButton := true else GS.ElipsisButton := false;

	stCellCur := GS.Cells[GS.col,GS.row];
	ErrorGs := false;
end;

procedure TOF_BTTVA.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if StCellCur <> GS.cells[Acol,Arow] then
  begin
    if not ControlDonnee (Acol,Arow) then BEGIN cancel := true; Exit; END;
    SetDonneeLigne (Acol,Arow);
  end;
end;

procedure TOF_BTTVA.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

end;

procedure TOF_BTTVA.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var Acol : integer;
begin
  if not ISDonneesOK (Ou,Acol) then
  BEGIN
//    cancel := true;
    GotoLigne (Ou,Acol);
  END;
end;

procedure TOF_BTTVA.GSEnter(Sender: TObject);
var cancel : boolean;
		Acol,Arow : integer;
begin
	ErrorGS := false;
	Arow := GS.row;
  Acol := GS.col;
  GSRowEnter(Self,Arow,cancel,false);
  GSCellEnter(Self,Acol,Arow,cancel);
end;

function TOF_BTTVA.ChercheCpte(var Cpte : string; Vide: Boolean): byte;
var St: string;
begin
  ChercheCpte := 0;
  St := UpperCase(cpte);
  Cache.Text := St;
  if ((Vide) and (St = '')) then exit;
  {$IFNDEF GCGC}
  {$IFNDEF COMPTAPAIE}
  if GChercheCompte(Cache, FicheGene) then
  {$ELSE}
  if GChercheCompte(Cache, nil) then
  {$ENDIF}
  {$ELSE}
  if GChercheCompte(Cache, nil) then
  {$ENDIF}
  begin
    if St <> Cache.Text then
    begin
      Cpte := Cache.Text;
    end;
    ChercheCpte := 2;
  end;
end;

procedure TOF_BTTVA.GSDblClick(Sender: TObject);
var theCpte : string;
begin
	if (GS.Col = CCPTEACH) or
  	 (GS.Col = CCPTEVTE) or
     (GS.Col = CENCAISACH) or
     (GS.Col = CENCAISVTE) or
     (GS.Col = CCPTEVTERG) then
  begin
    TheCpte := GS.cells[GS.col,GS.row];
    ChercheCpte(TheCpte, False);
    if TheCpte <> '' then if Length(TheCpte) < VH^.Cpta[fbGene].Lg then
    begin
      TheCpte := BourreLaDonc(TheCpte, fbGene);
    end;
    GS.Cells[GS.col,GS.row] := TheCpte; GS.InvalidateCell(GS.col,GS.row);
  end;
end;

procedure TOF_BTTVA.GSElipsisClick(Sender: TObject);
begin
  GSDBlCLick (sender);
end;

function TOF_BTTVA.ISDonneesOK (Arow : integer; var Acol : integer): boolean;
var Q: TQuery;
  	SQL: string;
begin
  Result := true;
  if (valeur(GS.Cells[CTAUXACH,Arow]) <> 0) and (GS.Cells[CCPTEACH,Arow] = '') then
  begin
    PgiError(TexteMessage[9], ecran.caption);
    result := false;
    Acol := CCPTEACH;
    Exit;
  end;
  if (valeur(GS.Cells[CTAUXVTE,Arow])<> 0) and (GS.Cells[CCPTEVTE,Arow] = '') then
  begin
    PgiError(TexteMessage[10], ecran.caption);
    result := false;
    Acol := CCPTEVTE;
    Exit;
  end;
  if GS.Cells[CCPTEACH,Arow] <> '' then
  begin
    SQL := FabReqTva (GS.Cells[CCPTEACH,Arow]);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      PgiError(TexteMessage[5], ecran.caption);
      result := false;
    end;
    Ferme(Q);
    if not result then
    begin
    	Acol := CCPTEACH;
    	Exit;
    end;
  end;
  if GS.Cells[CCPTEVTE,Arow] <> '' then
  begin
    SQL := FabReqTva(GS.Cells[CCPTEVTE,Arow]);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      PgiError(TexteMessage[6], ecran.caption);
      result := false;
    end;
    Ferme(Q);
    if not result then
    begin
    	Acol := CCPTEVTE;
    	Exit;
    end;
  end;
  if GS.Cells[CENCAISACH,Arow] <> '' then
  begin
    SQL := FabReqTva(GS.Cells[CENCAISACH,Arow]);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      PgiError(TexteMessage[13], ecran.caption);
      result := false;
    end;
    Ferme(Q);
    if not result then
    begin
    	Acol := CENCAISACH;
    	Exit;
    end;
  end;
  if GS.Cells[CENCAISVTE,Arow] <> '' then
  begin
    SQL := FabReqTva(GS.Cells[CENCAISVTE,Arow]);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      PgiError(TexteMessage[14], ecran.caption);
      result := false;
    end;
    Ferme(Q);
    if not result then
    begin
    	Acol := CENCAISVTE;
    	Exit;
    end;
  end;
  if GS.Cells[CCPTEVTERG,Arow] <> '' then
  begin
    if (not GetParamSocSecur('SO_GCVENTILRG', False)) or
       (not GetParamSocSecur('SO_RETGARANTIE', False)) then exit;
    //
    SQL := FabReqTva(GS.Cells[CCPTEVTERG,Arow]);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      PgiError(TexteMessage[16], ecran.caption);
      result := false;
    end;
    Ferme(Q);
    if not result then
    begin
    	Acol := CCPTEVTERG;
    	Exit;
    end;
  end;
end;

function TOF_BTTVA.ControlDonnee(Acol, Arow: integer): boolean;
var Q: TQuery;
  	SQL: string;
begin
	result := true;
  if (Acol = CCPTEACH) and  (valeur(GS.Cells[CCPTEACH,Arow]) <> 0) and (GS.Cells[CCPTEACH,Arow] = '') then
  begin
    PgiError(TexteMessage[9], ecran.caption);
    result := false;
    exit;
  end;
  if (Acol = CCPTEVTE)  and (valeur(GS.Cells[CCPTEVTE,Arow])<> 0) and (GS.Cells[CCPTEVTE,Arow] = '') then
  begin
    PgiError(TexteMessage[10], ecran.caption);
    result := false;
    exit;
  end;

  if (Acol = CCPTEACH) and (GS.Cells[CCPTEACH,Arow] <> '') then
  begin
    SQL := FabReqTva(GS.Cells[CCPTEACH,Arow]);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      PgiError(TexteMessage[5], ecran.caption);
      result := false;
    end;
    Ferme(Q);
  end else if (Acol = CCPTEVTE)  and (GS.Cells[CCPTEVTE,Arow] <> '') then
  begin
    SQL := FabReqTva(GS.Cells[CCPTEVTE,Arow]);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      PgiError(TexteMessage[6], ecran.caption);
      result := false;
    end;
    Ferme(Q);
  end else if (Acol = CENCAISACH) and (GS.Cells[CENCAISACH,Arow] <> '') then
  begin
    SQL := FabReqTva(GS.Cells[CENCAISACH,Arow]);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      PgiError(TexteMessage[13], ecran.caption);
      result := false;
    end;
    Ferme(Q);
  end else if (Acol=CENCAISVTE) and (GS.Cells[CENCAISVTE,Arow] <> '') then
  begin
    SQL := FabReqTva(GS.Cells[CENCAISVTE,Arow]);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      PgiError(TexteMessage[14], ecran.caption);
      result := false;
    end;
    Ferme(Q);
  end else if (Acol=CCPTEVTERG) and (GS.Cells[CCPTEVTERG,Arow] <> '') then
  begin
    SQL := FabReqTva(GS.Cells[CCPTEVTERG,Arow]);
    Q := OpenSql(SQL, True,-1,'',true);
    if Q.Eof then
    begin
      PgiError(TexteMessage[16], ecran.caption);
      result := false;
    end;
    Ferme(Q);
  end;
end;

procedure TOF_BTTVA.SetDonneeLigne(Acol,Arow: integer);
var TOBL : TOB;
begin
	if ARow < 1 then exit;
  TOBL := TOBTAXTVA.detail[Arow-1]; if TOBL = nil then exit;
  //
  if (Acol = CTAUXACH) then
  begin
  	TOBL.putValue('TV_TAUXACH',valeur(GS.cells[Acol,Arow]));
  end else if (Acol = CTAUXVTE) then
  begin
  	TOBL.putValue('TV_TAUXVTE',valeur(GS.cells[Acol,Arow]));
  end else if (Acol = CCPTEACH) then
  begin
  	TOBL.putValue('TV_CPTEACH',GS.cells[Acol,Arow]);
  end else if (Acol = CCPTEVTE) then
  begin
  	TOBL.putValue('TV_CPTEVTE',GS.cells[Acol,Arow]);
  end else if (Acol = CENCAISACH) then
  begin
  	TOBL.putValue('TV_ENCAISACH',GS.cells[Acol,Arow]);
  end else if (Acol=CENCAISVTE) then
  begin
  	TOBL.putValue('TV_ENCAISVTE',GS.cells[Acol,Arow]);
  end else if (Acol=CCPTEVTERG) then
  begin
  	TOBL.putValue('TV_CPTVTERG',GS.cells[Acol,Arow]);
  end else if (Acol=CCPTECOLLEC) then
  begin
  	TOBL.putValue('TV_COLLECTIF',GS.CellValues [Acol,Arow]);
  end;
end;

function TOF_BTTVA.FabReqTva(Cpt: string): string;
begin
  Result := 'Select G_GENERAL from GENERAUX Where G_GENERAL="' + Cpt + '"'
    + ' AND G_NATUREGENE<>"COC" And G_NATUREGENE<>"COD" And G_NATUREGENE<>"COF" And G_NATUREGENE<>"COS" '
    + ' AND G_COLLECTIF="-" AND G_NATUREGENE<>"IMO" AND G_NATUREGENE<>"CHA" AND G_NATUREGENE<>"PRO"';
end;

procedure TOF_BTTVA.GotoLigne(row, col: integer);
var Arow,Acol : integer;
		cancel : boolean;
begin
	SetEventsGS(false);
  Arow := Row;
  ACol := Col;
  GS.Col := Acol;
  GS.row := Arow;
  GSRowEnter(Self,Arow,cancel,false);
  GSCellEnter(Self,Acol,Arow,cancel);
	SetEventsGS(true);
  GS.ShowEditor;
end;

function TOF_BTTVA.ControleGrilleGS: boolean;
var Indice : integer;
		Colonne : integer;
begin
	result := true;
	for Indice := 0 to TOBTAXTVA.detail.count -1 do
  begin
  	if not ISDonneesOK (Indice+1,Colonne) then
    begin
    	result := false;
      GotoLigne (Indice+1,Colonne);
      break;
    end;
  end;
end;

function TOF_BTTVA.controleEntete : boolean;
var TheIndice : string;
begin
	result := true;
	if CC_CODE.text = '' then
  begin
  	PgiInfo ( TexteMessage[1], ecran.caption);
    CC_CODE.setFocus;
    result := false;
    exit;
  end;
	if TOBRTV.findFirst(['CC_CODE'],[CC_CODE.text],true) <> nil then
  begin
  	PgiInfo ( TexteMessage[3], ecran.caption);
    CC_CODE.setFocus;
    result := false;
    exit;
  end;
 if CC_LIBELLE.text = '' then
  begin
  	PgiInfo ( TexteMessage[18], ecran.caption);
    CC_LIBELLE.setFocus;
    result := false;
    exit;
  end;
  TheIndice := '0';
  if VALLIBRE.Value <> 0 then
  begin
    TheIndice := IntToSTr(VALLIBRE.Value);
    if TOBRTV.findFirst(['CC_LIBRE'],[TheIndice],true) <> nil then
    begin
      PgiInfo (TexteMessage[15], ecran.caption);
      VALLIBRE.Value := 0;
      VALLIBRE.setFocus;
      result := false;
      exit;
    end;
  end;
  CC_LIBRE.Text := TheIndice;

end;

procedure TOF_BTTVA.MajTXCPTA;
var CodeTaux : String;
		Sql : Hstring;
begin
  if TOBTAXTVA.IsOneModifie(true) then
  begin
    TOBTAXTVA.setAllModifie (true);
    CodeTaux := TOBTAXTVA.detail[0].getValue('TV_CODETAUX');
    // Suppression des données lignes
    Sql := 'DELETE FROM TXCPTTVA WHERE TV_TVAOUTPF="'+VH^.DefCatTVA+'" AND TV_CODETAUX="'+CodeTaux+'"';
    ExecuteSQL(Sql);
    // maj des donées lignes
    if not TOBTAXTVA.InsertDB(nil,false) then V_PGI.IoError := OeUnknown;
  end;
end;

function TOF_BTTVA.EcritDonnes : boolean;
var Indice : integer;
begin
	result := true;
  // cas de figure de la création d'une nouvelle définition
  if (TOBTAXTVA.detail[0].getValue('TV_TVAOUTPF')='') or (TOBTAXTVA.detail[0].getValue('TV_CODETAUX')='') then
  begin
    // -1- vérification des données en entete
    if not controleEntete then
    begin
      result := false;
      exit;
    end;
    // -2- maj des données entete dans les lignes
    for Indice := 0 to TOBTAXTVA.detail.count -1 do
    begin
      TOBTAXTVA.detail[Indice].putValue('TV_TVAOUTPF',VH^.DefCatTVA);
    	TOBTAXTVA.detail[Indice].PutValue('TV_CODETAUX',CC_CODE.text);
    end;
    TOBTAXTVA.SetAllModifie(true);
  end;
  if TRANSACTIONS(MajTXCPTA,0) <> OeOk then
  begin
  	PGIError('Problème durant la mise à jour');
  end else
  begin
  	TOBTAXTVA.setAllModifie (false);
  end;
end;

procedure TOF_BTTVA.GSExit(Sender: TObject);
var Arow,Acol : integer;
		cancel : boolean;
begin
	cancel := false;
  Arow := GS.row;
  ACol := GS.col;
  GSCellExit(Self,Acol,Arow,cancel);
  if cancel then ErrorGS := true else ErrorGS := false;
  if not cancel then GSRowExit(Self,Arow,cancel,false);
  if cancel then ErrorGS := true else ErrorGS := false;
end;

procedure TOF_BTTVA.BinsertClick(Sender: TObject);
var TOBL : TOB;
begin
  GSExit(self);
  if not ErrorGs then
  begin
    if TOBTAXTVA.IsOneModifie(true) then
    begin
      if not EcritDonnes then BEGIN {cancel := true;} EXIT; END;
    end;
  end;
  TOBL := TOB.create ('CHOIXCOD',TOBRTV,-1);
  TOBL.AddChampSupValeur('SEL','');
  TOBL.AddChampSupValeur('NEW','X');
  SetEventsRTV (false);
  AfficheGridRTV;
  SetEventsRTV (true);
  PositionneGRTV (TOBRTV.detail.count);
  PHAUT.enabled := true;
  SetEvents (true);
  CC_CODE.SetFocus;
  ModeCreat := true;
end;

procedure TOF_BTTVA.SetEvents(etat: boolean);
begin
	if etat then
  begin
    CC_CODE.onexit := CodeExit;
    CC_LIBELLE.onexit := LibelleExit;
    VALLIBRE.OnExit := VALLIBREExit;
  end else
  begin
    CC_CODE.onexit := nil;
    CC_LIBELLE.onexit := nil;
    VALLIBRE.OnExit := nil;
  end;
end;

procedure TOF_BTTVA.CodeExit(Sender: Tobject);
begin
	if CC_CODE.text = '' then
  begin
  	PgiInfo ( TexteMessage[1], ecran.caption);
    CC_CODE.setFocus;
    exit;
  end;
	if TOBRTV.findFirst(['CC_CODE'],[CC_CODE.text],true) <> nil then
  begin
  	PgiInfo ( TexteMessage[3], ecran.caption);
    CC_CODE.setFocus;
    exit;
  end;
end;

procedure TOF_BTTVA.LibelleExit(Sender: TObject);
begin
  if CC_LIBELLE.text = '' then
  begin
  	PgiInfo ( TexteMessage[18], ecran.caption);
    CC_LIBELLE.setFocus;
  end;
end;

procedure TOF_BTTVA.VALLIBREExit(Sender: TObject);
var TheIndice : String;
begin
  if VALLIBRE.Value <> 0 then
  begin
    TheIndice := IntToSTr(VALLIBRE.Value);
    if TOBRTV.findFirst(['CC_LIBRE'],[TheIndice],true) <> nil then
    begin
      PgiInfo (TexteMessage[15], ecran.caption);
      VALLIBRE.Value := 0;
      VALLIBRE.setFocus;
      exit;
    end;
  end;
  CC_LIBRE.Text := TheIndice;
end;

procedure TOF_BTTVA.BdeleteClick(Sender: Tobject);
var Q, QQ : TQuery;
		Sql, NomChoixCod : string;
begin
	if PgiAsk (TexteMessage[19], ecran.caption) = mryes then
  begin
  	Q := OpenSql('Select G_TVA From GENERAUX Where G_TVA="' + CC_CODE.text + '"', True,-1,'',true);
    if Q.eof then
    begin
      // Suppression des données lignes
      Sql := 'DELETE FROM TXCPTTVA WHERE TV_TVAOUTPF="'+VH^.DefCatTVA+'" AND TV_CODETAUX="'+CC_CODE.text+'"';
      ExecuteSQL(Sql);
      // Suppression des données entete

      // Récupération nom de base principale en cas de partage
      NomChoixcod := 'CHOIXCOD';
      QQ := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="TTTVA"', True);
      if not QQ.EOF then
      begin
        NomChoixCod := QQ.FindField('DS_NOMBASE').AsString+'.DBO.CHOIXCOD';
      end;
      Ferme (QQ);

      Sql := 'DELETE FROM '+NomChoixCod+' WHERE CC_TYPE="'+VH^.DefCatTVA+'" AND CC_CODE="'+CC_CODE.text+'"';
      ExecuteSQL(Sql);
      TOBRTV.detail[GRTV.row-1].free;
      GRTV.CacheEdit;
      GRTV.SynEnabled := False;
      GRTV.DeleteRow(GRTV.row);
//      GRTV.RowCount := GRTV.RowCount -1;
      GRTV.MontreEdit;
      GRTV.SynEnabled := true;
    end else
    begin
      PgiInfo ('Le code est référencé par un compte général.',ecran.caption);
    end;
    ferme (Q);
  end;

end;

procedure TOF_BTTVA.BimprimerClick(Sender: Tobject);
begin
	PrintDBGrid(GS,PHAUT,ecran.caption,'') 
end;

procedure TOF_BTTVA.BTVAETABClick(Sender: TObject);
begin
	ParamTVAEtab(CC_CODE.Text, VH^.DefCatTVA);
end;

Initialization
  registerclasses ( [ TOF_BTTVA ] ) ;
end.
