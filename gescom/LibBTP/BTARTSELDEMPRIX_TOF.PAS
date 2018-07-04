{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/02/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTARTSELDEMPRIX ()
Mots clefs ... : TOF;BTARTSELDEMPRIX
*****************************************************************}
Unit BTARTSELDEMPRIX_TOF ;

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
     graphics,
     ExtCtrls,
     sysutils,
     ComCtrls,
     HCtrls,
     Hpanel,
     HEnt1,
     HImgList,
     HMsgBox,
     UTOF,
     vierge,
     UEntCommun,
     HTB97 ;

Type
  TOF_BTARTSELDEMPRIX = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	TOBParam          : TOB;
    TOBPiece          : TOB;
    TOBPieceDemPrix   : TOB;
    TOBArticleDemPrix : TOB;
    TOBDetailDemprix  : TOB;
    TOBFournDemPrix   : TOB;
    TOBGestion        : TOB;
    //
    TheList           : THImageList;
    GS                : THgrid;

    CONFIGDEMANDEPRIX :  TToolbarButton97;
    BENREGSELECT      : TToolbarButton97;
    BCHERCHE          : TToolbarButton97;
    BSELECTALL        : TToolbarButton97;

    SELDEMANDEPRIX    : THValComboBox;
    TYPEARTICLE       : THValComboBox;
    FAMILLENIV1       : THValComboBox;
    FAMILLENIV2       : THValComboBox;
    FAMILLENIV3       : THValComboBox;
    //
    fFromTOB          : boolean;
    FlipFlop          : boolean;
    //
    fColNamesGS       : string;
    Falignement       : string;
    Ftitre            : string;
    fLargeur          : string;
    //
    procedure CreateTOB;
  	procedure DestroyObjects;
    procedure DestroyTOB;
    procedure DefinieGrid;
    procedure GetObjects;
    procedure SetInit;
    procedure DefiniGrille;
    procedure ChargeTOBgestion;
		procedure AfficheLagrille;
    procedure SetScreenEvents;
    function  EnregInSelect (TOBA : TOB) : boolean;
    //
    procedure BChercheClick (Sender : TObject);
    procedure BSelectAllClick (Sender : TObject);
    procedure BEnregSelectClick (Sender : TObject);
    procedure ConfigDemPrixClick (Sender :Tobject);
    procedure AffecteLIGDemPrix(TOBArtDemPrix,TOBDetailDemPrix,TOBFournDemPrix: TOB; SelUnique,newNumber: integer);
    procedure FamilleNiv1Change (Sender : Tobject);
    procedure FamilleNiv2Change (Sender : Tobject);
  end ;

Implementation

uses ImgList,UdemandePrix, Grids,UtilTOBPiece;

procedure TOF_BTARTSELDEMPRIX.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTARTSELDEMPRIX.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTARTSELDEMPRIX.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTARTSELDEMPRIX.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTARTSELDEMPRIX.OnArgument (S : String ) ;
begin
  Inherited ;
  fFromTOB := false;
  FlipFlop := false;

  if LaTOB <> nil then
  begin
    fFromTOB := true;
    TOBParam := laTOB;
    TOBArticleDemPrix := TOB (TOBParam.data);
    TOBDetailDemprix := TOB(TOBArticleDemPrix.Data);
    TOBFournDemPrix := TOB(TOBDetailDemprix.Data);
    TOBPieceDemPrix := TOB(TOBFournDemPrix.Data);
    TOBPiece        := TOB(TOBPieceDemPrix.Data);
  end;
  //
  CreateTOB;
  //
  GetObjects;
  SetInit;  // pour le combo box
  DefinieGrid; // définitioon des champs de la grille
  //
  ChargeTOBgestion;
  AfficheLagrille;
  SetScreenEvents;

end ;

procedure TOF_BTARTSELDEMPRIX.OnClose ;
begin
  DestroyTOB;
  DestroyObjects;
  Inherited ;
end ;

procedure TOF_BTARTSELDEMPRIX.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTARTSELDEMPRIX.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTARTSELDEMPRIX.CreateTOB;
begin
	TOBGestion := TOB.Create ('LES LIGNES',nil,-1);
  if (not fFromTOB) then
  begin
    TOBArticleDemPrix := TOB.create ('LES ARTDEMPRIX',nil,-1);
    TOBDetailDemprix := TOB.create ('LES DEtDEMPRIX',nil,-1);
    TOBFournDemPrix := TOB.create ('LES LIGFROUDEMPRIX',nil,-1);
    TOBPieceDemPrix := TOB.create ('LA PIECEDEMPRIX',nil,-1);
  end;
end;

procedure TOF_BTARTSELDEMPRIX.DestroyTOB;
begin
	TOBGestion.free;
  if (not fFromTOB) then
  begin
    TOBArticleDemPrix.free;
    TOBDetailDemprix.free;
    TOBFournDemPrix.free;
    TOBPieceDemPrix.free;
  end;
end;

procedure TOF_BTARTSELDEMPRIX.GetObjects;
begin
  GS := THgrid(GetControl('GS'));

  BCHERCHE          := TToolbarButton97 (GetControl('BCHERCHE'));
  CONFIGDEMANDEPRIX := TToolbarButton97 (GetControl('CONFIGDEMANDEPRIX'));
  BENREGSELECT      := TToolbarButton97 (GetControl('BENREGSELECT'));
  BSELECTALL        := TToolbarButton97 (getControl('BSELECTALL'));
  //
  SELDEMANDEPRIX    := THValComboBox(getControl('SELDEMANDEPRIX'));
  TYPEARTICLE       := THValComboBox (GetControl('TYPEARTICLE'));
  FAMILLENIV1       := THValComboBox (GetControl('FAMILLENIV1'));
  FAMILLENIV2       := THValComboBox (GetControl('FAMILLENIV2'));
  FAMILLENIV3       := THValComboBox (GetControl('FAMILLENIV3'));
  //
end;

procedure TOF_BTARTSELDEMPRIX.SetInit;
begin

  if (not fFromTOB) then
  begin
    SELDEMANDEPRIX.Plus := 'AND BPP_NATUREPIECEG="'+TOBparam.getString('NATUREPIECEG')+'" AND '+
                           'BPP_SOUCHE="'+TOBparam.getString('SOUCHE')+'" AND '+
                           'BPP_NUMERO='+IntToStr(TOBparam.getInteger('NUMERO'))+' AND '+
                           'BPP_INDICEG='+IntToStr(TOBparam.getInteger('INDICEG'));
  end else
  begin
    ChargeComboWithTob(SELDEMANDEPRIX, TobPieceDemPrix, 'BPP_LIBELLE', 'BPP_UNIQUE');
  end;

  //Chargement de Type d'Article pour bloquer uniquement MAR, ARP, PRE
  TYPEARTICLE.Plus := 'AND (CO_CODE="MAR" OR CO_CODE="ARP" OR CO_CODE="PRE")';

  FAMILLENIV2.Enabled := false;
  FAMILLENIV3.Enabled := false;

end;

procedure TOF_BTARTSELDEMPRIX.DefiniGrille;
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre,lelement : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep,Okimg : boolean;
    dec,NbCols : integer;
    indice,ind : integer;
    FFQTE,FFDEC : string;
begin
  //
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
  GS.ColCount := Nbcols;
  //
  st := fColNamesGS ;
  lesalignements := Falignement;
  lestitres := Ftitre;
  leslargeurs := fLargeur;

  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    GS.cells[Indice,0] := leTitre;
    GS.ColNames [Indice] := Nam;
    if copy(Alignement,1,1)='G' then GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then GS.ColAligns[indice] := taCenter;
    okImg := (copy(Alignement,8,1)='X');
    if OkVisu then
    begin
  		GS.ColWidths[indice] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
    end else
    begin
    	GS.ColWidths[indice] := -1;
    end;

    if (OkLib) or (okImg) then
    begin
    	GS.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	GS.ColDrawingModes[Indice]:= 'IMAGE';
      end;
    end else if (Dec<>0) or (Sep) then
    begin
      FFQTE := '#';
      if Dec > 0 then
      begin
        FFQTE := FFQTE+'0.';
        for ind := 1 to Dec do
        begin
          FFQTE := FFQTE + '0';
        end;
      end else FFQTE := FFQTE + '0';
    	if OkNulle then
      begin
        //
				FFDEC := FFQTE+';-'+FFQTE+'; ;';
      	GS.ColFormats[indice] := FFDEC ;
      end else
      begin
      	GS.ColFormats[indice] := FFQTE ;
      end;
    end;
  end ;
end;

procedure TOF_BTARTSELDEMPRIX.DefinieGrid;
begin
  // Définition de la liste de saisie pour la grille Détail
  fColNamesGS := 'SEL;LIBTYPEARTICLE;LIBFAMILLENIV1;LIBFAMILLENIV2;LIBFAMILLENIV3;CODEARTICLE;BDP_LIBELLE;';
  Falignement := 'G.0  ---;C.0  ---;G.0  ---;G.0  ---;G.0  ---;G.0  ---;G.0  ---;';
  Ftitre := ' ;Type;Famille 1;Famille 2;Famille 3;Article;Désignation;';
  fLargeur := '2;30;30;30;30;30;120;';
  //
  DefiniGrille;
end;

procedure TOF_BTARTSELDEMPRIX.ChargeTOBgestion;
var Indice : integer;
		TOBA,TOBB : TOB;
begin

  TOBgestion.ClearDetail;

	for Indice := 0 to TOBArticleDemPrix.detail.count -1 do
  begin
		TOBA := TOBArticleDemPrix.detail[Indice];
    if TOBA.getInteger('BDP_UNIQUE')=0 then
    begin
      if EnregInSelect (TOBA) then
      begin
        TOBB := TOB.Create ('ARTICLEDEMPRIX',TOBGestion,-1);
        TOBB.Dupliquer(TOBA,false,true);
        TOBB.Data := TOBA; // mémorise la tob provenance
      end;
    end;
  end;

  TOBGestion.Detail.Sort('TYPEARTICLE,FAMILLENIV1,FAMILLENIV2,FAMILLENIV3,CODEARTICLE');

end;

procedure TOF_BTARTSELDEMPRIX.AfficheLagrille;
var indice : integer;
begin
	GS.VidePile(false);
  if TOBGestion.detail.count = 0 then
		GS.RowCount := TOBgestion.detail.count +2
  else
		GS.RowCount := TOBgestion.detail.count +1;
    
  for Indice := 0 to TOBGestion.detail.count -1 do
  begin
  	TOBGestion.detail[Indice].PutLigneGrid (GS,Indice+1,false,true,fColNamesGS);
  end;
  TFVierge(ecran).HMTrad.ResizeGridColumns(GS);
  GS.row := 1;
end;

procedure TOF_BTARTSELDEMPRIX.DestroyObjects;
begin
	TheList.free;
end;

procedure TOF_BTARTSELDEMPRIX.SetScreenEvents;
begin
  BCHERCHE.onclick := BChercheClick;
  BSELECTALL.OnClick := BSelectAllClick;
  BENREGSELECT.onclick := BEnregSelectClick;
  CONFIGDEMANDEPRIX.onclick := ConfigDemPrixClick;
  FAMILLENIV1.OnChange := FamilleNiv1Change;
  FAMILLENIV2.OnChange := FamilleNiv2Change;

end;

procedure TOF_BTARTSELDEMPRIX.BChercheClick(Sender: TObject);
begin
	ChargeTOBgestion;
  AfficheLagrille;
  GS.row := 1;
end;

function TOF_BTARTSELDEMPRIX.EnregInSelect(TOBA: TOB): boolean;
begin
	result := true;
  if (TYPEARTICLE.value <> '') and (TOBA.GetString('TYPEARTICLE')<>TYPEARTICLE.Value) then
  begin
    result := false;
    exit;
  end;
  if (FAMILLENIV1.value <> '') and (TOBA.GetString('FAMILLENIV1')<>FAMILLENIV1.Value) then
  begin
    result := false;
    exit;
  end;
  if (FAMILLENIV2.value <> '') and (TOBA.GetString('FAMILLENIV2')<>FAMILLENIV2.Value) then
  begin
    result := false;
    exit;
  end;
  if (FAMILLENIV3.value <> '') and (TOBA.GetString('FAMILLENIV3')<>FAMILLENIV3.Value) then
  begin
    result := false;
    exit;
  end;
end;

procedure TOF_BTARTSELDEMPRIX.BSelectAllClick(Sender: TObject);
begin
	if GS.AllSelected = true then GS.AllSelected := false
  												 else GS.AllSelected := true;
end;

procedure TOF_BTARTSELDEMPRIX.AffecteLIGDemPrix (TOBArtDemPrix,TOBDetailDemPrix,TOBFournDemPrix: TOB; SelUnique,newNumber: integer);
var TOBD,TOBF : TOB;
		Indice : integer;
begin
	TOBD := TOBDetailDemprix.FindFirst(['UNIQUE','UNIQUELIG'],
  				[TOBARTDEMPRix.getInteger('BDP_UNIQUE'),TOBARTDEMPRix.getInteger('BDP_UNIQUELIG')],false);
  if TOBD <> nil then
  begin
    TOBD.SetInteger('UNIQUE',SelUnique);
    TOBD.SetInteger('UNIQUELIG',newNumber);
    for Indice := 0 to TOBD.detail.count -1 do
    begin
      TOBD.detail[indice].PutValue('BD0_UNIQUE',SelUnique);
      TOBD.detail[indice].PutValue('BD0_UNIQUELIG',newNumber);
    end;
  end;
	TOBF := TOBFournDemprix.FindFirst(['UNIQUE','UNIQUELIG'],
  				[TOBARTDEMPRix.getInteger('BDP_UNIQUE'),TOBARTDEMPRix.getInteger('BDP_UNIQUELIG')],false);
  if TOBF <> nil then
  begin
    TOBF.PutValue('UNIQUE',SelUnique);
    TOBF.PutValue('UNIQUELIG',newNumber);
    for Indice := 0 to TOBF.detail.count -1 do
    begin
      TOBF.detail[indice].PutValue('BD1_UNIQUE',SelUnique);
      TOBF.detail[indice].PutValue('BD1_UNIQUELIG',newNumber);
    end;
  end;
  TOBArtDemPrix.SetInteger ('BDP_UNIQUE',SelUnique);
  TOBArtDemPrix.SetInteger ('BDP_UNIQUELIG',newNumber);
end;

procedure TOF_BTARTSELDEMPRIX.BEnregSelectClick(Sender: TObject);
var Indice : integer;
		TOBG,TOBP : TOB;
    Unique,NewNum : integer;
    QQ : TQuery;
		Cledoc : r_cledoc;
begin
  if SELDEMANDEPRIX.Value='' then exit;
  NewNum := 0;
  Unique := StrToInt(SELDEMANDEPRIX.Value);
  TOBP   := TOBPieceDemPrix.findfirst(['BPP_UNIQUE'],[Unique],false);
  if TOBP = nil then
  begin
    cledoc.NaturePiece := TOBParam.getString('NATUREPIECEG');
    cledoc.Souche := TOBParam.getString('SOUCHE');
    cledoc.NumeroPiece := TOBParam.GetInteger('NUMERO');
    cledoc.Indice := TOBParam.GetInteger('INDICEG');
    QQ := OpenSql ('SELECT * FROM PIECEDEMPRIX WHERE '+
    				WherePiece(cledoc,TTdPieceDemPrix,false)+' AND BPP_UNIQUE="'+IntToStr(Unique)+'"',true,1,'',true);
    if not QQ.eof then
    begin
	    TOBP := TOB.create ('PIECEDEMPRIX',TOBPieceDemPrix,-1);
      TOBP.SelectDB('',QQ);
    end;
    ferme (QQ);
  end;
  if TOBP = nil then
  begin
    PgiError ('Impossible de charger la demande de prix ...');
    GS.AllSelected := false;
  	exit;
  end;

  for Indice := 1 to GS.RowCount do
  begin
    if GS.IsSelected(Indice) then
    begin
      TOBG := TOBGestion.detail[Indice -1];
      if TOBG.data = nil then continue;
      NewNum := GetNewUniqueLigDemprix (TOBP);
      AffecteLIGDemPrix (TOB(TOBG.data),TOBDetailDemprix,TOBFournDemprix,Unique,NewNum);
    end;
  end;
  GS.AllSelected := false;
  BChercheClick(self);
end;

procedure TOF_BTARTSELDEMPRIX.ConfigDemPrixClick(Sender: Tobject);
var cledoc : r_cledoc;
begin
  if FfromTob then
  begin
    // todo
    AfficheDemPrix (TOBPiece,TOBPieceDemPrix,TOBArticleDemPrix,TOBFournDemPrix,TOBDetailDemprix);
    ChargeComboWithTob(SELDEMANDEPRIX, TobPieceDemPrix, 'BPP_LIBELLE', 'BPP_UNIQUE');
   	ChargeTOBgestion;
    AfficheLagrille;
    GS.row := 1;
  end else
  begin
    cledoc.NaturePiece := TOBparam.getString('NATUREPIECEG');
    Cledoc.Souche :=  TOBparam.getString('SOUCHE');
    Cledoc.NumeroPiece := TOBparam.getInteger('NUMERO');
    cledoc.Indice := TOBparam.getInteger('INDICEG');
    AfficheDemprix (cledoc);
    //
    SELDEMANDEPRIX.ReLoad;
  end;

end;

procedure TOF_BTARTSELDEMPRIX.FamilleNiv1Change(Sender: Tobject);
begin
  FAMILLENIV2.value := '';
  FAMILLENIV3.value := '';
	if FAMILLENIV1.Value = '' then
  begin
    FAMILLENIV2.enabled := false;
    FAMILLENIV3.enabled := false;
  end else
  begin
		FAMILLENIV2.plus := 'and CC_LIBRE like "%'+FAMILLENIV1.Value+'%"';
		FAMILLENIV3.plus := 'and CC_LIBRE like "%'+FAMILLENIV1.Value+'%"';
    FAMILLENIV2.enabled := true;
    FAMILLENIV3.enabled := true;
  end;
end;

procedure TOF_BTARTSELDEMPRIX.FamilleNiv2Change(Sender: Tobject);
var st : string;
begin
  FAMILLENIV3.value := '';
	if FAMILLENIV2.Value = '' then
  begin
    FAMILLENIV3.enabled := false;
  end else
  begin
    st := 'and CC_LIBRE like "%'+FAMILLENIV1.Value;
    if FAMILLENIV2.value <> '' then st := st +FAMILLENIV2.Value;
		FAMILLENIV3.plus := st+'%"';
    FAMILLENIV3.enabled := true;
  end;

end;

Initialization
  registerclasses ( [ TOF_BTARTSELDEMPRIX ] ) ;
end.
