{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 25/09/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREAJUSTEANA ()
Mots clefs ... : TOF;BTREAJUSTEANA
*****************************************************************}
Unit BTREAJUSTEANA_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     Fe_Main,
{$else}
     eMul, 
{$ENDIF}
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HPanel,
     HMsgBox, 
     UTOF,
     uEntCommun,
     UtilTOBPiece,
     Aglinit,
     FactTOB,
     utofAfBaseCodeAffaire,
     vierge,
     HTB97,
     Types,
     Windows,
     Paramsoc,
     Messages,
     Graphics,
     Grids,
     SaisUtil
     ;

const
  ListeChamps = '_ESP;REFERENCE;LIBELLE;QTE;UNITE;MONTANT;FAMILLENIV2;SUPR';
  largeur = '2;17;50;12;4;12;25;4;';
  alignement = 'G.0O X--;G.0O X--;G.0O X--;D/2  ---;C.0O ---;D/2  ---;G.0 $---;G.0O X--;';
  Titre = ' ;Référence;Désignation;Qte;Unité;Montant;Analytique; ;';
  NC = '0;0;0;0;0;0;0;0;'; //definition des zones saisissable ou non

  ListeSaisie = 'BLR_INDDETAIL;BLR_FAMILLENIV2;BLR_QTE;BLR_MONTANTACH;SUPR';
  LArgSaisie = '3;25;12;12;4;';
  aligsaisie = 'C.0O ---;G.0O X--;D/2  ---;D/2  ---;G.0O X--;';
  TitreSaisie = ' ;Code analytique;Qte;Montant;Suppr.;';
  NCSaisie = '0;1;1;1;0'; //definition des zones saisissable ou non

Type

  TTyperesult = (TTRCode,TTRLibelle);

  TRefAna = record
  	Entity : Integer;
    Journal : string;
    Exercice : String;
    Axe : string;
    Piece : integer;
    Ligne : Integer;
  end;


  TOF_BTREAJUSTEANA = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    IFAMILLE,IQTE,IMONTANT : Integer;
    ComptaOn : boolean;
    AxeVentil : string;
    DEV : Rdevise;
    DateMvt : TDateTime;
    fListeChamps : string;
    fListeSaisie : string;
    TOBREAJ,TOBREAJ_O,TOBPiece,TOBSaisie,TOBConso,TOBECRANA,TOBArticles: TOB;
    TOBTiers : TOB;
    TOBAffaire : TOB;
    TOBCpta : TOB;
    TOBStructAna,TOBStructAnalib : TOB;
    ECARTLIGNE : ThNumEdit;
    TOTALREAJUSTE,ECART : ThNumEdit;
    GS : THGrid;
    GS2 : THGrid;
    stCell : string;
    JnlOd : string;
    BTSUPLIGNE,BTREAJUSTE,BTSUPLIGNEDET : TToolbarButton97;
    procedure ChargeReajustements;
    procedure ConstitueSaisie;
    procedure AddChampsSupSaisie;
    procedure SetInfoPiece;
    procedure DefiniAttributCol(GS: THgrid; Nom: string; Colonne: integer;LaLargeur, Lalignement, Letitre, LeNC: string);
    procedure ConstitueAffichage;
    procedure AfficheLaGrille (GS : Thgrid; TOBLignes : TOB;  fListe : string; Depart : integer =0) ;
    procedure AfficheLaLigne(GS: THGrid; TOBLignes: TOB; Ligne: integer; Fliste : string);
    procedure AddLigneSaisie(TOBL: TOB);
    //
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    //
    procedure GS2RowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GS2RowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GS2CellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GS2CellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    //
    procedure AfficheBoutonsSup(Etat : boolean);
    procedure ShowDroite(Etat : boolean);
    procedure PositionneBoutons(TOBL :TOB;Ou : integer);
    procedure ChargeModifAnal(TOBL : TOB);
    procedure ConstitueAffichageGS2 (TOBL : TOB);
    procedure SupLigneClick (sender : TObject);
    procedure ReajusteLigneClick (Sender : TObject);
    procedure AjouteMontantLigne(TOBD : TOB; Sens : string='+');
    procedure PositionneDeleteLigne (TOBL : TOB; Ou : Integer);
    procedure SUPLigDetailClick (Sender : TObject);
    procedure ActiveEventsGs2(Etat : Boolean=false);
    procedure ZoneSuivanteOuOk(var ACol, ARow: Integer;var Cancel: boolean);
    function ZoneAccessible(ACol, ARow: Integer): boolean;
    procedure InsertLigne;
    procedure NettoieLigneVide(TOBLIGNE: TOB);
    procedure BdeleteALLClick (Sender : Tobject);
    procedure BValiderClick (Sender : Tobject);
    procedure PrepareVentilationAnalytique;

    procedure SupprimeAnciennesDatas;
    procedure EcritODConsommations;
    procedure EcritDatasAnalytique;
    procedure AjouteConsommation(TOBD,TOBLigne : TOB);

    procedure GS2KeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    function DemandeDatesMvts(var DateFac: TDateTime): boolean;
    procedure GSGetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    function ExistDatas : boolean;
    procedure ClearGrilleGs2;
    function RecupSection(TOBATrait: TOB; TTR: TTyperesult;AFFAIREREF: TOB): string;
    procedure ChargeStructureAna;
    function EncodeRefAna(TOBTT : TOB) : string;
    procedure DecodeRefAna(Chaine : string ; var RefAna : TRefAna);
    procedure SoldeLigne(Arow : integer);

  public
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
  end ;

procedure ReajusteAnalytique (TheChaine : string);

Implementation
uses utilPGI,Ent1,EntGC,CalcOLEGenericBTP,AffaireUtil,USpecifPoc,UtilSaisieConso,FactCpta,VentilCpta;

procedure ReajusteAnalytique (TheChaine : string);
var CleDoc      : R_CleDoc;
  StA           : string;
  QQ : TQuery;
  TOBPiece : TOB;
  JnlOd : string;
begin
  JnlOd := GetParamSocSecur ('SO_BTREPARTANALACH','');
  if JnlOd = '' then
  begin
    PgiError ('Le journal des OD analytiques n''est pas défini');
    Exit;
  end;
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  TOBPIece := TOB.Create('PIECE',nil,-1);
  TRY
    StA := TheChaine;
    StringToCleDoc(StA, CleDoc);
    //
    QQ := OpenSql ('SELECT * FROM PIECE WHERE '+WherePiece(cledoc,ttdPiece,False),True,1,'',True);
    if not QQ.eof then
    begin
      TOBPiece.SelectDB('',QQ);
    end;
    ferme (QQ);
    if TOBPiece.GetString('GP_NATUREPIECEG')<> 'FF' then Exit;
    QQ := OpenSql ('SELECT * FROM LIGNE WHERE '+WherePiece(cledoc,ttdligne,false),True,-1,'',True);
    if not QQ.eof then
    begin
      TOBPiece.loaddetailDB('LIGNE','','',QQ,false);
    end;
    Ferme(QQ);
    //
    TheTOB := TOBpiece;
    AGLLanceFiche('BTP','BTREAJUSTEANA','','','ACTION=MODIFICATION');
  FINALLY
    TheTOB := nil;
    TOBPiece.free;
  END;
end;

procedure TOF_BTREAJUSTEANA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREAJUSTEANA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREAJUSTEANA.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTREAJUSTEANA.OnLoad ;
var cancel : boolean;
begin
  Inherited ;
  //
  cancel := false;
  AfficheLaGrille(GS,TOBSaisie,fListeChamps);
  TFVierge(ecran).HMTrad.ResizeGridColumns(GS);
  BTSUPLIGNE.onclick := SupLigneClick;
  BTREAJUSTE.onclick := ReajusteLigneClick;
  BTSUPLIGNEDET.OnClick := SUPLigDetailClick;
  GS.OnRowEnter := GSRowEnter;
  GS.OnRowExit := GSRowExit;
  GS.GetCellCanvas := GSGetCellCanvas;
  GSRowEnter(Self,1,cancel,false);
  JnlOd := GetParamSocSecur ('SO_BTREPARTANALACH','');

  if JnlOd = '' then
  begin
    PgiError ('Le journal des OD analytiques n''est pas défini');
    TToolbarButton97 (GetControl('Bvalider')).enabled := false;
  end;
  if AxeVentil = '' then
  begin
    PgiError ('Vauillez paramétrer la modification analytique achat');
    TToolbarButton97 (GetControl('Bvalider')).enabled := false;
  end;
  if TOBStructAna.detail.Count = 0 then
  begin
    PgiError ('Vauillez définir la ventilation analytique achat');
    TToolbarButton97 (GetControl('Bvalider')).enabled := false;
  end;

end ;

procedure TOF_BTREAJUSTEANA.ChargeStructureAna;
var SQL : string;
    QQ : TQuery;
begin
  SQL := 'SELECT * '+
  			 'FROM DECOUPEANA '+
         'WHERE GDA_AXE="'+AxeVentil+'" AND GDA_TYPECOMPTE="ACH" AND GDA_TYPESTRUCTANA="SEC" '+
  			 'ORDER BY GDA_RANG';
  QQ := OpenSql (SQL,True,-1,'',True);
  if not QQ.Eof then
  begin
  	TOBStructAna.LoadDetailDB('DECOUPEANA','','',QQ,false);
  end;
  ferme (QQ);
  //
  SQL := 'SELECT * '+
  			 'FROM DECOUPEANA '+
         'WHERE GDA_AXE="'+AxeVentil+'" AND GDA_TYPECOMPTE="ACH" AND GDA_TYPESTRUCTANA="LSE" '+
  			 'ORDER BY GDA_RANG';
  QQ := OpenSql (SQL,True,-1,'',True);
  if not QQ.Eof then
  begin
  	TOBStructAnalib.LoadDetailDB('DECOUPEANA','','',QQ,false);
  end;
  Ferme(QQ);
end;

procedure TOF_BTREAJUSTEANA.OnArgument (S : String ) ;
begin
  Inherited ;
  AxeVentil := GetParamSocSecur('SO_BTAXEANALACH','');
  //
  ReinitTOBAffaires;
  TOBStructAna := TOB.Create('LA STRUCTURE ANA',nil,-1);
  TOBStructAnalib := TOB.Create('LES LIB ANA',nil,-1);
  TOBCpta := TOB.Create('LES VENTILS',nil,-1);
  TOBArticles := TOB.Create ('LES ARTICLEs',nil,-1);
  TOBConso := TOB.Create ('LES CONSO',nil,-1);
  TOBECRANA := TOB.Create('LES ECR',nil,-1);
  TOBTiers := TOB.Create('TIERS',nil,-1);
  TOBREAJ := TOB.Create('LES REAJUSTEMENTS',nil,-1);
  TOBREAJ_O := TOB.Create ('LES ANCIENS',nil,-1);
  TOBSaisie := TOB.Create('LA SAISIE',nil,-1);
  AddChampsSupSaisie;
  TOBPiece := LaTOB;
  DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.GetDouble('GP_TAUXDEV');
  //
  ComptaOn := (GetInfoParPiece(TOBPiece.getString('GP_NATUREPIECEG'),'GPP_TYPEECRCPTA')='NOR');
  StockeCetteAffaire (TOBPiece.getString('GP_AFFAIRE'));
  TOBAffaire := FindCetteAffaire (TOBPiece.getString('GP_AFFAIRE'));
  ChargeStructureAna;
  //
  ChargeReajustements;
  ConstitueSaisie;
  //
  GS := THGrid(GetControl('GS'));
  GS2 := THGrid(GetControl('GS2'));
  BTSUPLIGNE :=  TToolBarButton97(GetControl('BTSUPLIGNE'));
  BTREAJUSTE :=  TToolBarButton97(GetControl('BTREAJUSTE'));
  BTSUPLIGNEDET := TToolBarButton97(GetControl('BTSUPLIGNEDET'));
  ECARTLIGNE := ThNumEdit (GeTControl('ECARTLIGNE'));
  TOTALREAJUSTE := ThNumEdit (GeTControl('TOTALREAJUSTE'));
  ECART := ThNumEdit (GeTControl('ECART'));
  TToolbarButton97 (GetControl('Bdelete')).OnClick := BdeleteALLClick;
  TToolbarButton97 (GetControl('BVALIDER')).OnClick := BValiderClick;

  ConstitueAffichage;

  TOBSaisie.PutEcran(Ecran,THPanel(Getcontrol('PHAUT')));
  TOBSaisie.PutEcran(Ecran,THPanel(Getcontrol('PBAS')));
end ;

procedure TOF_BTREAJUSTEANA.OnClose ;
begin
  ReinitTOBAffaires;
  TOBConso.free;
  TOBREAJ.free;
  TOBREAJ_O.free;
  TOBSaisie.free;
  TOBECRANA.Free;
  TOBTiers.Free;
  TOBArticles.Free;
  TOBStructAna.free;
  TOBStructAnalib.free;
  Inherited ;
end ;

procedure TOF_BTREAJUSTEANA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREAJUSTEANA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTREAJUSTEANA.ChargeReajustements;
var QQ : TQuery;
    SQL : string;
    cledoc : r_cledoc;
begin
  SQL := 'SELECT * FROM TIERS WHERE T_TIERS="'+TOBPiece.GetString('GP_TIERS')+'"';
  QQ := OpenSQL(SQL,True,1,'',true);
  TOBTiers.SelectDB('',QQ);
  ferme (QQ);
  //
  Cledoc:=TOB2CleDoc(TOBPiece);
  SQL := 'SELECT * FROM LIGREAANAL WHERE '+WherePiece(Cledoc,ttdReajAnal,false)+' ORDER BY BLR_NUMORDRE, BLR_INDDETAIL';
  QQ := OpenSql (SQL,True,-1,'',true);
  if not QQ.eof then
  begin
    TOBREAJ.LoadDetailDB('LIGREAANAL','','',QQ,false);
  end;
  ferme (QQ);
  TOBREAJ_O.Dupliquer(TOBREAJ,True,true);
end;

procedure TOF_BTREAJUSTEANA.ConstitueSaisie;
var II : Integer;
    TOBL : TOB;
    TOBD : TOB;
    TOBR : TOB;
begin
  TOBSaisie.ClearDetail; TOBSaisie.InitValeurs(false);
  SetInfoPiece;
  for II := 0 to TOBPiece.Detail.count -1 do
  begin
    TOBL := TOBPiece.detail[II];
    TOBD := TOB.Create ('ONE LIG',TOBSaisie,-1);
    AddLigneSaisie(TOBD);
    TOBD.SetString('TYPELIGNE',TOBL.getString('GL_TYPELIGNE'));
    TOBD.SetInteger('NUMLIGNE',TOBL.getInteger('GL_NUMLIGNE'));
    TOBD.SetInteger('NUMORDRE',TOBL.getInteger('GL_NUMORDRE'));
    TOBD.SetString('REFERENCE',TOBL.getString('GL_CODEARTICLE'));
    TOBD.SetString('LIBELLE',TOBL.getString('GL_LIBELLE'));
    TOBD.SetDouble('QTE',TOBL.GetDouble('GL_QTEFACT'));
    TOBD.SetString('UNITE',TOBL.GetString('GL_QUALIFQTEACH'));
    TOBD.SetDouble('MONTANT',TOBL.GetDouble('GL_MONTANTHTDEV'));
    TOBD.SetString('FAMILLENIV1',TOBL.GetString('GL_FAMILLENIV1'));
    TOBD.SetString('FAMILLENIV2',TOBL.GetString('GL_FAMILLENIV2'));
    TOBD.Data := TOBL;
  end;
  TOBSaisie.SetDouble('TOTALREAJUSTE',TOBPiece.GetDouble('GP_TOTALHTDEV'));
  TOBSaisie.SetDouble('ECART',0);

  if TOBReaj.detail.count = 0 then exit;
  II := 0;
  repeat
    TOBR := TOBREAJ.detail[II];
    TOBD := TOBSaisie.FindFirst(['NUMORDRE'],[TOBR.GetInteger('BLR_NUMORDRE')],true);
    if TOBD <> nil then
    begin
      TOBR.ChangeParent(TOBD,-1);
      TOBSaisie.SetDouble('TOTALREAJUSTE',TOBSaisie.GetDouble('TOTALREAJUSTE')+TOBR.GetDouble('BLR_MONTANTACH'));
      TOBD.SetString('REAFFECT','X');
    end else inc(II);
  until II >= TOBREAJ.detail.Count;
  TOBSaisie.SetDouble('ECART',TOBPiece.GetDouble('GP_TOTALHTDEV')-TOBSaisie.GetDouble('TOTALREAJUSTE'));
end;

procedure TOF_BTREAJUSTEANA.AddLigneSaisie (TOBL : TOB);
begin
  TOBL.AddChampSupValeur('_ESP','');
  TOBL.AddChampSupValeur('TYPELIGNE','');
  TOBL.AddChampSupValeur('NUMLIGNE',0);
  TOBL.AddChampSupValeur('NUMORDRE',0);
  TOBL.AddChampSupValeur('REFERENCE','');
  TOBL.AddChampSupValeur('LIBELLE','');
  TOBL.AddChampSupValeur('QTE',0.00);
  TOBL.AddChampSupValeur('UNITE','');
  TOBL.AddChampSupValeur('MONTANT',0.00);
  TOBL.AddChampSupValeur('FAMILLENIV1','');
  TOBL.AddChampSupValeur('FAMILLENIV2','');
  TOBL.AddChampSupValeur('REAFFECT','-');
end;

procedure TOF_BTREAJUSTEANA.AddChampsSupSaisie;
begin
  TOBSaisie.AddChampSupValeur('GP_LIBNATUREPIECE','');
  TOBSaisie.AddChampSupValeur('GP_NUMERO',0);
  TOBSaisie.AddChampSupValeur('GP_DATEPIECE',iDate1900);
  TOBSaisie.AddChampSupValeur('GP_TIERS',''); TOBSaisie.AddChampSupValeur('LIBTIERS','');
  TOBSaisie.AddChampSupValeur('GP_AFFAIRE','');
  TOBSaisie.AddChampSupValeur('GP_AFFAIRE0','');
  TOBSaisie.AddChampSupValeur('GP_AFFAIRE1','');
  TOBSaisie.AddChampSupValeur('GP_AFFAIRE2','');
  TOBSaisie.AddChampSupValeur('GP_AFFAIRE3','');
  TOBSaisie.AddChampSupValeur('GP_AVENANT','');
  TOBSaisie.AddChampSupValeur('GP_REFINTERNE','');
  TOBSaisie.AddChampSupValeur('GP_REFEXTERNE','');
  TOBSaisie.AddChampSupValeur('GP_BSVREF','');
  TOBSaisie.AddChampSupValeur('GP_TOTALHTDEV',0);
  TOBSaisie.AddChampSupValeur('TOTALREAJUSTE',0);
  TOBSaisie.AddChampSupValeur('ECART',0);
end;


procedure TOF_BTREAJUSTEANA.DefiniAttributCol(GS : THgrid; Nom: string; Colonne: integer; LaLargeur,Lalignement, Letitre, LeNC: string);
var Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep : boolean;
    dec : integer;
    FF : string;
begin

  TransAlign(Lalignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
  GS.cells[Colonne,0] := leTitre;
  // Alignement dans la colonne
  if copy(LAlignement,1,1)='G' then GS.ColAligns[Colonne] := taLeftJustify
  else if copy(LAlignement,1,1)='D' then GS.ColAligns[Colonne] := taRightJustify
  else if copy(LAlignement,1,1)='C' then GS.ColAligns[Colonne] := taCenter;
  GS.ColWidths[Colonne] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
  GS.Collengths[Colonne] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
  if OkLib then
  begin
    GS.ColFormats[Colonne] := 'CB=' + Get_Join(Nom);
  end else if (Dec<>0) or (Sep) then
  begin
    GS.ColFormats[Colonne] := FF+';-'+FF+'; ;' ;
  end;
  if LeNc = '0' then GS.ColEditables [Colonne]:= false;
end;


procedure TOF_BTREAJUSTEANA.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,Tiers_: THEdit);
begin
  inherited;
  Aff  := THEdit(GetControl('GP_AFFAIRE'));
  Aff0 := THEdit(GetControl('GP_AFFAIRE0'));
  Aff1 := THEdit(GetControl('GP_AFFAIRE1'));
  Aff2 := THEdit(GetControl('GP_AFFAIRE2'));
  Aff3 := THEdit(GetControl('GP_AFFAIRE3'));
  Aff4 := THEdit(GetControl('GP_AVENANT'));
end;


procedure TOF_BTREAJUSTEANA.AfficheLaLigne (GS : THGrid; TOBLignes : TOB; Ligne : integer; Fliste : string);
var TOBL : TOB;
begin
  TOBL := TOBlignes.detail[Ligne-1]; if TOBl = nil then exit;
  TOBL.PutLigneGrid (GS,Ligne,false,false,Fliste);
  GS.InvalidateRow(Ligne);
end;


procedure TOF_BTREAJUSTEANA.AfficheLaGrille (GS : Thgrid; TOBLignes : TOB; fListe : string; Depart : integer =0) ;
var Indice : integer;
begin
  GS.RowCount := TOBLignes.detail.count+1; if GS.rowCount < 2 then GS.rowCount := 2;
  For Indice := Depart to TOBLignes.detail.count -1 do
  begin
    AfficheLaLigne (GS, TOBLignes, Indice+1,fListe);
  end;
  GS.Invalidate;
end;

procedure TOF_BTREAJUSTEANA.ClearGrilleGs2;
var II,JJ : Integer;
begin
  for II :=0 to GS.RowCount -1 do
  begin
    for JJ := 0 to GS.ColCount -1 do
    begin
      GS2.cells[JJ,II] := '';
    end;
  end;
end;


procedure TOF_BTREAJUSTEANA.ConstitueAffichageGS2 (TOBL : TOB);
var lesElements,lesLargeurs,lesAlignements,lesTitres,lesNC : string;
    Lelement,LaLargeur,Lalignement,LeTItre,LeNC : string;
    Colonne : integer;
begin
  GS2.VidePile(false);
  GS2.RowCount := TOBL.detail.count +2;
  ClearGrilleGs2;
  lesElements := ListeSaisie;
  lesLargeurs := LArgSaisie;
  lesAlignements := aligsaisie;
  lesTitres := TitreSaisie;
  lesNc := NCSaisie;
  //
  Lelement := ReadtokenSt(lesElements);
  LaLargeur := ReadtokenSt(lesLargeurs);
  Lalignement := ReadtokenSt(lesAlignements);
  LeTItre := readTokenSt(lesTitres);
  LeNC := readTokenSt(lesNc);
  Colonne := 0;
  GS2.colCount := 5;

  fListeSaisie := ListeSaisie;

  Repeat
    if Lelement = '' then break;
    DefiniAttributCol (GS2,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    if Lelement='BLR_FAMILLENIV2' then
    begin
      GS2.ColFormats[colonne] := 'CB=GCFAMILLENIV2| AND CC_LIBRE="'+TOBL.getString('FAMILLENIV1')+'"|<<Tous>>';
      IFAMILLE := Colonne;
    end else if Lelement='BLR_QTE' then
    begin
      IQTE := Colonne;
    end else if Lelement='BLR_MONTANTACH' then
    begin
      IMONTANT := Colonne;
    end else if lelement = 'SUPR' then
    begin
      GS2.colLengths[colonne] := -1;
    end;
    //
    Lelement := ReadtokenSt(lesElements);
    LaLargeur := ReadtokenSt(lesLargeurs);
    Lalignement := ReadtokenSt(lesAlignements);
    LeTItre := readTokenSt(lesTitres);
    LeNC := readTokenSt(lesNc);
    inc(colonne);
    //
  until lelement = '';
end;



procedure TOF_BTREAJUSTEANA.ConstitueAffichage;
var lesElements,lesLargeurs,lesAlignements,lesTitres,lesNC : string;
    Lelement,LaLargeur,Lalignement,LeTItre,LeNC : string;
    Colonne : integer;
begin
  GS.VidePile(false);
  lesElements := ListeChamps;
  lesLargeurs := Largeur;
  lesAlignements := alignement;
  lesTitres := titre;
  lesNc := NC;
  //
  Lelement := ReadtokenSt(lesElements);
  LaLargeur := ReadtokenSt(lesLargeurs);
  Lalignement := ReadtokenSt(lesAlignements);
  LeTItre := readTokenSt(lesTitres);
  LeNC := readTokenSt(lesNc);
  Colonne := 0;
  GS.colCount := 8;

  fListeChamps := ListeChamps;

  Repeat
    if Lelement = '' then break;
    DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    //
    Lelement := ReadtokenSt(lesElements);
    LaLargeur := ReadtokenSt(lesLargeurs);
    Lalignement := ReadtokenSt(lesAlignements);
    LeTItre := readTokenSt(lesTitres);
    LeNC := readTokenSt(lesNc);
    inc(colonne);
    //
  until lelement = '';
end;


procedure TOF_BTREAJUSTEANA.SetInfoPiece;
begin
  TOBSaisie.SetString('GP_LIBNATUREPIECE',RechDom('GCNATUREPIECEG',TOBPiece.GetString('Gp_NATUREPIECEG'),false));
  TOBSaisie.SetInteger('GP_NUMERO',TOBPiece.GetInteger('GP_NUMERO'));
  TOBSaisie.SetDateTime('GP_DATEPIECE',TOBPiece.GetDateTime('GP_DATEPIECE'));
  TOBSaisie.SetString('GP_TIERS',TOBPiece.GetString('GP_TIERS'));
  TOBSaisie.SetString('LIBTIERS',RechDom('GCTIERSSAISIE', TOBPiece.GetString('GP_TIERS'), False));
  TOBSaisie.SetString('GP_AFFAIRE',TOBPiece.GetString('GP_AFFAIRE'));
  TOBSaisie.SetString('GP_AFFAIRE0',Copy(TOBPiece.GetString('GP_AFFAIRE'),1,1));
  TOBSaisie.SetString('GP_AFFAIRE1',TOBPiece.GetString('GP_AFFAIRE1'));
  TOBSaisie.SetString('GP_AFFAIRE2',TOBPiece.GetString('GP_AFFAIRE2'));
  TOBSaisie.SetString('GP_AFFAIRE3',TOBPiece.GetString('GP_AFFAIRE3'));
  TOBSaisie.SetString('GP_AVENANT',TOBPiece.GetString('GP_AVENANT'));
  TOBSaisie.SetString('GP_REFINTERNE',TOBPiece.GetString('GP_REFINTERNE'));
  TOBSaisie.SetString('GP_REFEXTERNE',TOBPiece.GetString('GP_REFEXTERNE'));
  TOBSaisie.SetString('GP_BSVREF',TOBPiece.GetString('GP_BSVREF'));
  TOBSaisie.SetDouble('GP_TOTALHTDEV',TOBPiece.getdouble('GP_TOTALHTDEV'));
  TOBSaisie.SetDouble('TOTALREAJUSTE',TOBPiece.getdouble('GP_TOTALHTDEV'));
end;

procedure TOF_BTREAJUSTEANA.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
begin
  TOBL := TOBSaisie.detail[Ou-1]; if TOBl = nil then exit;
  AfficheBoutonsSup(false);
  ShowDroite(false);
  if TOBL.getString('TYPELIGNE')<>'ART' then exit;
  PositionneBoutons(TOBL,Ou);
  ChargeModifAnal(TOBL);
end;

procedure TOF_BTREAJUSTEANA.NettoieLigneVide(TOBLIGNE : TOB);
var II : Integer;
    TOBD : TOB;
begin
  if TOBLigne.detail.count = 0 then exit; 
  II := 0;
  repeat
    TOBD := TOBLigne.detail[II];
    if (TOBD.getString('BLR_FAMILLENIV2')='') then TOBD.Free else inc(II);
  until II = TOBLIGNE.detail.Count;
end;

procedure TOF_BTREAJUSTEANA.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TOBLIGNE : TOB;
begin
  ECARTLIGNE.Value := ARRONDI(ECARTLIGNE.Value,DEV.Decimale);
  if ECARTLIGNE.Value <> 0 then
  begin
    PGIInfo('IMPOSSIBLE : L''écart entre le montant facturé et le montant réajusté n''est pas nul.');
    cancel := True;
    exit;
  end;
  TOBLigne := TOBSaisie.detail[Ou-1]; if TOBligne = nil then exit;
  NettoieLigneVide(TOBLIGNE);
  if (GS2.Visible) and (GS2.Col = 1) then
  begin
    ActiveEventsGs2 (false);
    SendMessage(GS2.Handle, WM_KeyDown, VK_TAB, 0);
  end;
  GS2.RowCount := TOBLIGNE.detail.count +2; 
end;

procedure TOF_BTREAJUSTEANA.AfficheBoutonsSup(Etat: boolean);
begin
  BTSUPLIGNE.visible := Etat;
  BTREAJUSTE.visible := Etat;
end;

procedure TOF_BTREAJUSTEANA.ActiveEventsGs2 (Etat : Boolean=false);
begin
  if not Etat then
  begin
    GS2.OnrowEnter := nil;
    GS2.OnrowExit := nil;
    GS2.OnCellEnter := nil;
    GS2.OnCellExit := nil;
    GS2.OnKeyDown := nil;
  end else
  begin
    GS2.OnrowEnter := GS2RowEnter;
    GS2.OnrowExit := GS2Rowexit;
    GS2.OnCellEnter := GS2CellEnter;
    GS2.OnCellExit := GS2CellExit;
    GS2.OnKeyDown := Gs2KeyDown;
  end;
end;

procedure TOF_BTREAJUSTEANA.GS2KeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
//var Vide, Cancel: Boolean;
begin
//  Vide := (Shift = []);
//  Cancel := False;
  if Key = 123 then
  begin
    Key := 0;
    SoldeLigne(GS2.row);
  end;

end;


procedure TOF_BTREAJUSTEANA.ShowDroite(Etat: boolean);
begin
  GS2.CacheEdit; GS.HideEditor;
  GS2.Visible := Etat;
  ActiveEventsGs2 (False);
  THPanel(GetControl('PECARTLIG')).Visible := Etat;
  if etat then ActiveEventsGs2 (true);
end;

procedure TOF_BTREAJUSTEANA.PositionneBoutons(TOBL: TOB; Ou: integer);
var  Arect: Trect;
begin
  ARect := GS.CellRect(7, Ou);
  if TOBL.Detail.count > 0 then
  begin
    with BTSUPLIGNE do
    begin
      Top := Arect.top - GS.GridLineWidth;
      Left := Arect.Left;
      Width := Arect.Right - Arect.Left;
      Height := Arect.Bottom - Arect.Top;
      Parent := GS;
      visible := True;
    end;
  end else
  begin
    with BTREAJUSTE do
    begin
      Top := Arect.top - GS.GridLineWidth;
      Left := Arect.Left;
      Width := Arect.Right - Arect.Left;
      Height := Arect.Bottom - Arect.Top;
      Parent := GS;
      visible := True;
    end;
  end;
end;

procedure TOF_BTREAJUSTEANA.ChargeModifAnal(TOBL: TOB);
var cancel : boolean;
    Acol,Arow : Integer;
begin
  if TOBL.detail.count > 0 then
  begin
    ConstitueAffichageGS2 (TOBL);

    AfficheLaGrille(GS2,TOBL,fListeSaisie );
    GS2.RowCount := TOBL.detail.count +2;

    TFVierge(ecran).HMTrad.ResizeGridColumns(GS2);
    ShowDroite(true);
    Arow := 1;
    Acol := 1;
    GS2RowEnter (Self,1,cancel,false);
    GS2CellEnter (Self,Acol,Arow,cancel);
  end;
end;

procedure TOF_BTREAJUSTEANA.GS2CellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBLIGNE : TOB;
begin
  TOBLigne := TOBSaisie.detail[GS.row-1]; if TOBligne = nil then exit;
  ZoneSuivanteOuOk(ACol, ARow, Cancel);
  if Cancel then Exit;
  stCell := GS2.cells[Acol,Arow];
end;

procedure TOF_BTREAJUSTEANA.GS2CellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBLigne : TOB;
    TOBD : TOB;
begin
  if (Acol = IFAMILLE) and (GS2.CellValues [ACol,Arow]='') and (GS2.Row >= Arow) then BEGIN Cancel := True; Exit; end
  else if (Acol = IQTE) and (valeur(GS2.Cells[ACol,Arow])=0.0) and (GS2.Row >= Arow)  then BEGIN cancel := true; exit; END
  else if (Acol = IMONTANT) and (valeur(GS2.Cells[ACol,Arow])=0.0) and (GS2.Row >= Arow)  then BEGIN cancel := true; exit; END;

  TOBLigne := TOBSaisie.detail[GS.row-1]; if TOBligne = nil then exit;
  TOBD := TOBLigne.detail[Arow-1]; if TOBD = nil then exit;
  //
  if (ACol=IFAMILLE) then
  begin
    if TOBD.getInteger('BLR_INDDETAIL')=1 then BEGIN GS2.Cells [ACol,Arow]:= StCell; Exit; END;
    TOBD.SetString('BLR_FAMILLENIV2',GS2.CellValues [ACol,Arow]);
  end else if (ACol=IQTE) then
  begin
    if TOBD.getInteger('BLR_INDDETAIL')=1 then BEGIN GS2.Cells [ACol,Arow]:= StCell; Exit; END;
    AjouteMontantLigne(TOBD,'-');
    TOBD.SetDouble('BLR_QTE',VALEUR(GS2.Cells [ACol,Arow]));
    TOBD.SetDouble('BLR_MONTANTACH',ARRONDI(TOBD.GetDouble('BLR_QTE')*TOBD.GetDouble('BLR_DPA'),DEV.Decimale));
    AjouteMontantLigne(TOBD);
  end else if (ACol=IMONTANT) then
  begin
    if TOBD.getInteger('BLR_INDDETAIL')=1 then BEGIN GS2.Cells [ACol,Arow]:= StCell; Exit; END;
    AjouteMontantLigne(TOBD,'-');
    TOBD.SetDouble('BLR_MONTANTACH',VALEUR(GS2.Cells [ACol,Arow]));
    TOBD.SetDouble('BLR_QTE',ARRONDI(TOBD.GetDouble('BLR_MONTANTACH')/TOBD.GetDouble('BLR_DPA'),V_PGI.OkDecQ));
    AjouteMontantLigne(TOBD);
  end;
  AfficheLaLigne (GS2,TOBLigne,ARow,fListeSaisie);
end;

procedure TOF_BTREAJUSTEANA.GS2RowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TOBLigne : TOB;
    TOBD : TOB;
begin
  TOBLigne := TOBSaisie.detail[GS.row-1]; if TOBligne = nil then exit;
  if Ou > TOBLigne.detail.count then BEGIN InsertLigne; GS2.RowCount := TOBLigne.detail.count +2; END;
  TOBD := TOBLigne.detail[Ou-1]; if TOBD = nil then exit;
  PositionneDeleteLigne(TOBD,Ou);
end;

procedure TOF_BTREAJUSTEANA.GS2RowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TOBLIGNE : TOB;
begin
  TOBLigne := TOBSaisie.detail[GS.row-1]; if TOBligne = nil then exit;
  NettoieLigneVide(TOBLIGNE);
  GS2.RowCount := TOBLigne.detail.count +2;
  cancel := false;
end;

procedure TOF_BTREAJUSTEANA.ReajusteLigneClick(Sender: TObject);
var TOBLigne,TOBL : TOB;
    TOBD : TOB;
    cancel : Boolean;
begin
  cancel := false;
  TOBLigne := TOBSaisie.detail[GS.row-1]; if TOBligne = nil then exit;
  TOBL := TOB(TOBLigne.Data);
  TOBD := TOB.Create('LIGREAANAL',TOBLigne,-1);
  TOBD.SetString('BLR_NATUREPIECEG',TOBL.GetString('GL_NATUREPIECEG'));
  TOBD.SetString('BLR_SOUCHE',TOBL.GetString('GL_SOUCHE'));
  TOBD.SetInteger('BLR_NUMERO',TOBL.GetInteger('GL_NUMERO'));
  TOBD.SetInteger('BLR_INDICEG',TOBL.GetInteger('GL_INDICEG'));
  TOBD.SetInteger('BLR_NUMORDRE',TOBL.GetInteger('GL_NUMORDRE'));
  TOBD.SetInteger('BLR_INDDETAIL',1);
  TOBD.SetString('BLR_FAMILLENIV1',TOBL.GetString('GL_FAMILLENIV1'));
  TOBD.SetString('BLR_FAMILLENIV2',TOBL.GetString('GL_FAMILLENIV2'));
  TOBD.SetString('BLR_FAMILLENIV3',TOBL.GetString('GL_FAMILLENIV3'));
  TOBD.Setdouble('BLR_NUMMOUV',0);
  TOBD.SetDouble('BLR_MONTANTACH',TOBL.GetDouble('GL_TOTALHTDEV')*-1);
  TOBD.SetDouble('BLR_QTE',TOBL.GetDouble('GL_QTEFACT')*-1);
  TOBD.SetDouble('BLR_DPA',TOBL.GetDouble('GL_PUHTDEV'));
  AjouteMontantLigne(TOBD);
  GSRowEnter(Self,GS.Row,cancel,false);
end;

procedure TOF_BTREAJUSTEANA.SupLigneClick(sender: TObject);
var II : Integer;
    TOBLIgne,TOBD : TOB;
    cancel : Boolean;
begin
  if PGIAsk('ATTENTION : Vous allez supprimer la modification analytique de cette ligne.#13#10 Confirmez-vous?')<>mryes then exit;
  TOBLigne := TOBSaisie.detail[GS.row-1]; if TOBligne = nil then exit;
  for II := 0 to TOBLIgne.Detail.Count -1 do
  begin
    TOBD := TOBLIgne.detail[II];
    AjouteMontantLigne(TOBD,'-');
  end;
  TOBLIgne.ClearDetail;
  cancel := False;
  if (GS2.Visible) and (GS2.Col = 1) then
  begin
    ActiveEventsGs2 (false);
    SendMessage(GS2.Handle, WM_KeyDown, VK_TAB, 0);
  end;
  //
  GSRowEnter(Self,GS.Row,cancel,false);
  GS.Invalidate;
end;

procedure TOF_BTREAJUSTEANA.AjouteMontantLigne(TOBD: TOB; Sens : string='+');
begin
  if Sens = '+' then
  begin
    ECARTLIGNE.Value := ECARTLIGNE.Value + TOBD.GetDouble('BLR_MONTANTACH');
    TOTALREAJUSTE.Value := TOTALREAJUSTE.Value + TOBD.GetDouble('BLR_MONTANTACH');
    ECART.Value := TOBSaisie.GetDouble('GP_TOTALHTDEV')-TOTALREAJUSTE.Value;
  end else
  begin
    ECARTLIGNE.Value := ECARTLIGNE.Value - TOBD.GetDouble('BLR_MONTANTACH');
    TOTALREAJUSTE.Value := TOTALREAJUSTE.Value - TOBD.GetDouble('BLR_MONTANTACH');
    ECART.Value := TOBSaisie.GetDouble('GP_TOTALHTDEV')-TOTALREAJUSTE.Value;
  end;
end;

procedure TOF_BTREAJUSTEANA.PositionneDeleteLigne(TOBL: TOB; Ou: Integer);
var  Arect: Trect;
begin
  ARect := GS2.CellRect(4, Ou);
  with BTSUPLIGNEDET do
  begin
    Top := Arect.top - GS.GridLineWidth;
    Left := Arect.Left;
    Width := Arect.Right - Arect.Left;
    Height := Arect.Bottom - Arect.Top;
    Parent := GS2;
    visible := false;
    if TOBL.GetInteger('BLR_INDDETAIL')<>1 then visible := true;
  end;
end;

procedure TOF_BTREAJUSTEANA.SUPLigDetailClick(Sender: TObject);

  procedure ReindiceTOBD(TOBligne : TOB);
  var II : Integer;
  begin
    for II := 0 to TOBligne.detail.count -1 do
    begin
      TOBligne.detail[II].SetInteger('BLR_INDDETAIL',II+1);
    end;
  end;

var TOBLigne : TOB;
    TOBD : TOB;
    cancel : Boolean;
begin
  cancel := falSe;
  if PGIAsk('ATTENTION : Vous allez supprimer cette ligne.#13#10 Confirmez-vous ?')<> Mryes then exit;
  TOBLigne := TOBSaisie.detail[GS.row-1]; if TOBligne = nil then exit;
  TOBD := TOBLigne.detail[Gs2.row-1]; if TOBD = nil then exit;
  AjouteMontantLigne (TOBD,'-');
  TOBD.Free;
  ReindiceTOBD(TOBligne);
  ActiveEventsGs2(false);
  GS2.DeleteRow(Gs2.row);
  ActiveEventsGs2(true);
  GS2RowEnter(Self,GS2.Row,cancel,false);
  GS.Invalidate;
end;

procedure TOF_BTREAJUSTEANA.InsertLigne;
var TOBLIGNE,TOBL,TOBD : TOB;
begin
  TOBLigne := TOBSaisie.detail[GS.row-1]; if TOBligne = nil then exit;
  TOBL := TOB(TOBLigne.Data);
  TOBD := TOB.Create('LIGREAANAL',TOBLigne,-1);
  TOBD.SetString('BLR_NATUREPIECEG',TOBL.GetString('GL_NATUREPIECEG'));
  TOBD.SetString('BLR_SOUCHE',TOBL.GetString('GL_SOUCHE'));
  TOBD.SetInteger('BLR_NUMERO',TOBL.GetInteger('GL_NUMERO'));
  TOBD.SetInteger('BLR_INDICEG',TOBL.GetInteger('GL_INDICEG'));
  TOBD.SetInteger('BLR_NUMORDRE',TOBL.GetInteger('GL_NUMORDRE'));
  TOBD.SetInteger('BLR_INDDETAIL',TOBLIgne.detail.count);
  TOBD.SetString('BLR_FAMILLENIV1',TOBL.GetString('GL_FAMILLENIV1'));
  TOBD.SetString('BLR_FAMILLENIV2','');
  TOBD.SetString('BLR_FAMILLENIV3','');
  TOBD.Setdouble('BLR_NUMMOUV',0);
  TOBD.SetDouble('BLR_QTE',0);
  TOBD.SetDouble('BLR_DPA',TOBL.GetDouble('GL_PUHTDEV'));
  TOBD.SetDouble('BLR_MONTANTACH',Arrondi(TOBD.GetDouble('BLR_QTE')*TOBD.GetDouble('BLR_DPA'),DEV.decimale));
end;


procedure TOF_BTREAJUSTEANA.ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
  RowFirst : integer;
begin
	RowFirst := ARow;
  OldEna := GS2.SynEnabled;
  GS2.SynEnabled := False;
  Sens := -1;
  ChgLig := (GS2.Row <> ARow);
  ChgSens := False;
  if GS2.Row > ARow then Sens := 1 else if ((GS2.Row = ARow) and (ACol <= GS2.Col)) then Sens := 1;
  ACol := GS2.Col;
  ARow := GS2.Row;
  ii := 0;
  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      Lim := GS2.RowCount - 1;
      // ---
      if ((ACol = GS2.ColCount - 1) and (ARow >= Lim)) then
      begin
        if (ChgSens) then Break else
        begin
          InsertLigne;
          break;
        end;
      end;
      if ChgLig then
      begin
        ACol := GS2.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GS2.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := GS2.FixedCols;
      end;
    end else
    begin
      if ((ACol = GS2.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := GS2.ColCount;
        ChgLig := False;
      end;
      if ACol > GS2.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GS2.ColCount - 1;
      end;
    end;
  end;
  GS2.SynEnabled := OldEna;
  if (Arow <> RowFirst) then
  BEGIN
  	GS2RowExit (self,RowFirst,cancel,false);
  	if Cancel then
    begin
    	GS2RowEnter (self,RowFirst,cancel,false);
    end else
    begin
    	GS2RowEnter (self,Arow,cancel,false);
    end;
  END;
end;


function TOF_BTREAJUSTEANA.ZoneAccessible(ACol, ARow: Longint): boolean; { NEWPIECE }
begin
  Result := True;
end;

procedure TOF_BTREAJUSTEANA.BdeleteALLClick(Sender: Tobject);
var II : Integer;
    cancel : Boolean;
begin
  if PGIAsk('ATTENTION : Vous allez supprimer TOUTE la répartition analytique de cette facture.#13#10 Confirmez-vous ?')<> Mryes then exit;
  cancel := false;
  if (GS2.Visible) and (GS2.Col = 1) then
  begin
    ActiveEventsGs2 (false);
    SendMessage(GS2.Handle, WM_KeyDown, VK_TAB, 0);
  end;
  for II := 0 to TOBSaisie.detail.count -1 do
  begin
    if TOBSaisie.Detail[II].detail.Count > 0 then TOBSaisie.Detail[II].ClearDetail;
  end;
  TOTALREAJUSTE.Value := TOBPiece.GetDouble('GP_TOTALHTDEV');
  ECART.Value := 0;
  GSRowEnter(Self,GS.row,cancel,false);
end;


function TOF_BTREAJUSTEANA.DemandeDatesMvts(var DateFac: TDateTime ) : boolean;
var TobDates : TOB;
begin
  TOBDates := TOB.Create ('LES DATES', nil,-1);
  TOBDates.AddChampSupValeur('RETOUROK','-');
  TOBDates.AddChampSupValeur('DATFAC',iDate1900);
  TOBDates.AddChampSupValeur('DATESITUATION','-');
  TOBDates.AddChampSupValeur('TYPEDATE','Date de mise à jour');
  TRY
    TheTOB := TOBDates;
    AGLLanceFiche('BTP','BTDEMANDEDATES','','','');
    TheTOB := nil;
    if TOBDates.getValue('RETOUROK')='X' then
    begin
    	DateFac := TOBDates.GetDateTime('DATFAC');
    end;
  FINALLY
  	result := (TOBDates.getValue('RETOUROK')='X');
  	freeAndNil(TOBDates);
  END;
end;

procedure TOF_BTREAJUSTEANA.BValiderClick(Sender: Tobject);
begin
  ECART.Value := ARRONDI(ECART.Value,DEV.Decimale);
  if ECART.Value <> 0 then
  begin
    PGIInfo('IMPOSSIBLE : L''écart entre le montant facturé et le montant réajusté n''est pas nul.');
    Ecran.ModalResult := mrNone;
    Exit;
  end;
  if (ExistDatas) and (not DemandeDatesMvts (Datemvt)) then
  begin
    Ecran.ModalResult := mrNone;
    Exit;
  end;

  BEGINTRANS;
  TRY
    SupprimeAnciennesDatas;
    if ComptaOn then PrepareVentilationAnalytique;
    EcritODConsommations;
    if ComptaOn then EcritDatasAnalytique;
    COMMITTRANS;
  except
    PGIInfo('IMPOSSIBLE : Erreur dans l''écriture des OD''s');
    ROLLBACK;
  end;
end;


function TOF_BTREAJUSTEANA.RecupSection(TOBATrait: TOB; TTR : TTyperesult; AFFAIREREF : TOB ): string;
var II : Integer;
		TOBS : TOB;
    Lng : Integer;
    Champs,CodeVide,value,CodeAna : string;
begin
	Result := '';
  CodeVide := '';
  if TTR = TTRCode then
  begin
    for II:=0 to TOBStructAna.Detail.Count-1 do
    begin
      TOBS := TOBStructAna.detail[II];
      Lng := TOBS.GetInteger('GDA_LONGUEUR');
      Champs := TOBS.GetString('GDA_LIBCHAMP');
      if TOBATrait.FieldExists(Champs) then
      begin
        value := Copy(TOBATrait.GetString(Champs),1,Lng);
      end else
      begin
        if AFFAIREREF.FieldExists(Champs) then value := copy(AFFAIREREF.GetString(Champs),1,Lng);
      end;
      CodeVide := CodeVide + BourreChamps ('',Lng,TFichierBase(Ord(fbAxe1)),True);
      Value:=BourreChamps (Value,Lng,TFichierBase(Ord(fbAxe1)),True);
      CodeAna := CodeAna + value;
    end;
    if length(CodeAna)>0 then Result := BourreLaDonc(CodeAna,TFichierBase(Ord(fbAxe1)))
    else Result := '';
  end else
  begin
    for II:=0 to TOBStructAnalib .Detail.Count-1 do
    begin
      TOBS := TOBStructAnalib.detail[II];
      Lng := TOBS.GetInteger('GDA_LONGUEUR');
      Champs := TOBS.GetString('GDA_LIBCHAMP');
      if TOBATrait.FieldExists(Champs) then
      begin
        value := Trim(Copy(TOBATrait.GetString(Champs),1,Lng));
      end else
      begin
        if AFFAIREREF.FieldExists(Champs) then value := Trim(Copy(AFFAIREREF.GetString(Champs),1,Lng));
      end;
      CodeAna := CodeAna + value;
    end;
    Result := Trim(CodeAna);
  end;
end;


procedure TOF_BTREAJUSTEANA.PrepareVentilationAnalytique;

  function findCpta(FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe,VenteAchat : string) : TOB;
  var TOBC,TOBXX : TOB;
      LastSql,SQL : string;
      Q : TQuery;
  begin
    LastSQL:='' ; SQL:='' ;
    TOBC:=FindTOBCode(TOBCpta,FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe) ;
    if TOBC=Nil then
    BEGIN
      SQL:=FabricSQLCompta(FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe,VenteAchat) ;
      Q:=OpenSQL(SQL,True,-1, '', True) ;
      if Not Q.EOF then
      BEGIN
        TOBXX := TOB.Create ('LES VENTIL',nil,-1);
        TOBXX.LoadDetailDB('CODECPTA','','',Q,false);
        TOBC:=FindTOBCode(TOBXX,FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe) ;
        if TOBC <> nil then
        begin
          TOB.Create('Ecr',TOBC,-1) ;
          TOB.Create('Stk',TOBC,-1) ;
          TOBC.ChangeParent(TOBCpta,-1);
        end;
        TOBXX.Free;
      END ;
      Ferme(Q);
    end;
    if TOBC = nil then
    begin
      TOBC := CreerTOBCodeCpta (TOBCpta);
      TOBC.SetString('GCP_COMPTAARTICLE',FamArt);
      TOBC.SetString('GCP_COMPTATIERS',FamTiers);
      TOBC.SetString('GCP_COMPTAAFFAIRE',FamAff);
      TOBC.SetString('GCP_REGIMETAXE',Regime);
      TOBC.SetString('GCP_FAMILLETAXE',FamTaxe);
      TOBC.SetString('GCP_ETABLISSEMENT',Etab);
      TOBC.SetString('GCP_VENTEACHAT',VenteAchat);
      TOBC.SetString('GCP_CPTEGENEACH',VH_GC.GCCpteHTACH);
      TOBC.SetString('GCP_CPTEGENEVTE',VH_GC.GCCpteHTVTE);
    end;
    Result := TOBC;
    TOBCpta.Detail.Sort('-GCP_COMPTAARTICLE;-GCP_COMPTATIERS;-GCP_COMPTAAFFAIRE;-GCP_REGIMETAXE;-GCP_FAMILLETAXE;-GCP_ETABLISSEMENT');
  end;


  procedure AjouteOdanal (TOBl,TOBD : TOB; CompteGene,Section,Libelle : string; Numcpta : integer);
  var TTF1 : TOB;
      MtMVt : double;
  begin
    TTF1 := TOB.Create('ANALYTIQ',TOBECRANA,-1);
    //
    if TOBD.GetIndex = 0 then TTF1.SetString('Y_TYPEMVT','AE')
                         else TTF1.SetString('Y_TYPEMVT','AL');
    TTF1.SetInteger ('Y_NUMVENTIL',TOBD.GetIndex+1);
    //
    TTF1.SetDateTime('Y_DATECOMPTABLE',DateMvt);
    TTF1.SetInteger ('Y_PERIODE',GetPeriode(DateMvt)) ;
    TTF1.SetInteger ('Y_SEMAINE',NumSemaine(DateMvt)) ;
    TTF1.SetInteger ('Y_NUMEROPIECE',NumCpta);
    TTF1.SetInteger ('Y_NUMLIGNE',0);
    //
    TTF1.setString  ('Y_GENERAL',CompteGene);
    TTF1.setString  ('Y_AXE',AxeVentil);
    TTF1.setString  ('Y_SECTION',Section);
    TTF1.setString  ('Y_EXERCICE',QUELEXODT(DateMvt));
    //
    TTF1.SetString  ('Y_LIBELLE',Libelle);
    TTF1.SetString  ('Y_NATUREPIECE','OD');
    TTF1.SetString  ('Y_QUALIFPIECE','N');
    TTF1.SetString  ('Y_TYPEANALYTIQUE','X');
    TTF1.SetString  ('Y_ETAT','0000000000');
    TTF1.SetString  ('Y_UTILISATEUR',V_PGI.User);
    TTF1.SetDateTime('Y_DATECREATION',DateMvt);
    TTF1.SetString  ('Y_SOCIETE',V_PGI.CodeSociete);
    TTF1.SetString  ('Y_ETABLISSEMENT',TOBl.GetString('GL_ETABLISSEMENT'));
    TTF1.SetString  ('Y_DEVISE'            , V_PGI.DevisePivot ) ;
    TTF1.SetDouble  ('Y_TAUXDEV'           , 1 ) ;
    TTF1.SetDateTime('Y_DATETAUXDEV'       , DateMvt ) ;
    TTF1.SetString  ('Y_CONTROLE'          , '-' ) ;
    TTF1.SetString  ('Y_QUALIFQTE1'        , '...' ) ;
    TTF1.SetString  ('Y_QUALIFQTE2'        , '...' ) ;
    TTF1.SetString  ('Y_QUALIFECRQTE1'     , '...' ) ;
    TTF1.SetString  ('Y_QUALIFECRQTE2'     , '...' ) ;
    TTF1.SetString  ('Y_ECRANOUVEAU'       , 'N' ) ;
    TTF1.SetString  ('Y_CREERPAR'          , 'SAI' ) ;
    TTF1.SetString  ('Y_EXPORTE'           , '---' ) ;
    TTF1.SetString  ('Y_JOURNAL'           , JNLOD ) ;
    TTF1.SetString  ('Y_VALIDE'            , '-' ) ;
    TTF1.SetString  ('Y_CONFIDENTIEL'      , '0' ) ;
    TTF1.SetDateTime('Y_DATPER',iDate1900) ;
    TTF1.SetInteger ('Y_ENTITY',0) ;
    TTF1.SetString  ('Y_REFGUID','') ;
    MtMVt := TOBD.GetDouble('BLR_MONTANTACH');
    if MtMvt >= 0 then
    begin
      TTF1.SetDouble('Y_DEBITDEV',TTF1.GetDouble('Y_DEBITDEV')+Abs(MtMvt));
      TTF1.SetDouble('Y_DEBIT',TTF1.GetDouble('Y_DEBIT')+ DEVISETOPIVOT(Abs(MtMvt),DEV.Taux,DEV.Quotite));
    end else
    begin
      TTF1.SetDouble('Y_CREDITDEV',TTF1.GetDouble('Y_CREDITDEV')+Abs(MtMvt));
      TTF1.SetDouble('Y_CREDIT',TTF1.GetDouble('Y_CREDIT')+DEVISETOPIVOT(Abs(MtMvt),DEV.Taux,DEV.Quotite));
    end;
    TOBD.SetString('BLR_REFANALYTIQUE',EncodeRefAna(TTF1));
  end;

  procedure TraiteAnalLigne (TOBLIG,TOBD : TOB; Numcpta : integer);
  var CompteGene,regime,Etab,Famtiers,FamArt,FamAff,Nature,VenteAchat,NatV,FamTaxe,Section,LibSection : string;
      TOBC,TOBA,TOBL : TOB;
      QQ : TQuery;
  begin
    TOBL := TOB(TOBLIG.data);
    TOBA := TOBArticles.FindFirst(['GA_ARTICLE'],[TOBL.GetString('GL_ARTICLE')],true);
    if TOBA = nil then
    begin
      QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TOBL.GetString('GL_ARTICLE')+'"',True,1,'',true);
      if not QQ.eof then
      begin
        TOBA := TOB.Create('ARTICLE',TOBArticles,-1);
        TOBA.selectdb ('',QQ);
      end;
      ferme (QQ);
      if TOBA = nil then
      begin
        Exception.create('Article '+TOBL.GetString('GL_ARTICLE')+' innexistant');
        Exit;
      end;
    end;
    Regime:=TOBPiece.GetValue('GP_REGIMETAXE');
    Etab:=TOBPiece.GetValue('GP_ETABLISSEMENT');
    FamTiers:=TOBTiers.GetValue('T_COMPTATIERS');
    FamArt:=TOBA.GetValue('GA_COMPTAARTICLE') ;
    FamAff := TraiteFamilleAffaire (TOBL.GetValue('GL_AFFAIRE'), TobAffaire );
    FamTaxe := TOBL.GetValue('GL_FAMILLETAXE1');
    //
    Nature:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
    VenteAchat:=GetInfoParPiece(Nature,'GPP_VENTEACHAT') ;
    if VenteAchat='VEN' then NatV:='HV' else NatV:='HA' ;
    TOBC  := findCpta(FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe,VenteAchat);
    if TOBC = nil then
    begin
      Exception.Create ('Compte comptable non défini');
      Exit;
    end;
    CompteGene := TOBC.GetString('GCP_CPTEGENEACH'); // c'est une facture d'achat
    //
    TOBL.SetString('GL_FAMILLENIV2',TOBD.GetString('BLR_FAMILLENIV2'));
    Section := RecupSection (TOBL,TTRCode,TOBAffaire);
    if Section = '' then
    begin
      Exception.Create ('Section '+Section+' impossible a insérer');
      Exit;
    end;
    LibSection := RecupSection(TOBL,TTRLibelle,TOBAffaire);
    CreerSectionVolee ( Section,LibSection,AxeVentil );
    AjouteODANal (TOBL,TOBD,CompteGene,Section,LibSection,NumCpta);
  end;

var II,JJ : Integer;
    TOBL : TOB;
    NumCpta : Integer;
begin
  for II := 0 to TOBSaisie.detail.count - 1 do
  begin
    TOBL := TOBSaisie.detail[II];
    if TOBL.detail.count = 0 then Continue;
    NumCpta:=GetNewNumJal(JnlOd,true,DateMvt); // pour chaque compte général

    for JJ := 0 to TOBL.Detail.count -1 do
    begin
      TRY
        if (TOBL.detail[JJ].getString('BLR_FAMILLENIV2')='') or (TOBL.detail[JJ].GetDouble('BLR_MONTANTACH')=0) then continue;
        // 
	      TraiteAnalLigne (TOBL,TOBL.detail[JJ],NumCpta);
      EXCEPT
        raise;
      end;
    end;
  end;
end;

procedure TOF_BTREAJUSTEANA.EcritDatasAnalytique;
begin
  TOBECRANA.InsertDB(nil);
end;

procedure TOF_BTREAJUSTEANA.EcritODConsommations;
var ii,JJ : Integer;
    TOBD,TOBS,TOBL : TOB;
begin
  TOBREAJ.ClearDetail;
  for II := 0 to TOBSaisie.Detail.Count -1 do
  begin
    TOBS :=TOBSaisie.Detail[II];
    if TOBS.detail.count = 0 then Continue;
    TOBL := TOB(TOBS.data);
    JJ := 0;
    repeat
      TOBD := TOBS.detail[JJ];
      if (TOBD.GetString('BLR_FAMILLENIV2')='') or (TOBD.Getdouble('BLR_MONTANTACH')=0.0) then
      begin
        Inc(JJ);
        Continue;
      end;
      TOBD.SetAllModifie(true);
      AjouteConsommation(TOBD,TOBL);
      TOBD.ChangeParent(TOBREAJ,-1); 
    until JJ >= TOBS.detail.count ;
  end;
  TOBREAJ.InsertDB(nil);
  TOBConso.InsertDB(nil); 
end;

procedure TOF_BTREAJUSTEANA.SupprimeAnciennesDatas;

  procedure  DeleteEcrAna (RefAna : TRefAna);
  begin
    ExecuteSQL('DELETE FROM ANALYTIQ WHERE '+
                 'Y_ENTITY='+InttoStr(RefAna.Entity)+' AND '+
                 'Y_JOURNAL="'+RefAna.Journal+'" AND '+
                 'Y_EXERCICE="'+RefAna.Exercice+'" AND '+
                 'Y_AXE="'+RefAna.Axe+'" AND '+
                 'Y_NUMEROPIECE='+InttoStr(RefAna.Piece)+' AND '+
                 'Y_NUMLIGNE='+InttoStr(RefAna.Ligne));
  end;

var cledoc : R_CLEDOC;
    II : Integer;
    TOBR: TOB;
    RANA : TRefAna;
begin
  Cledoc:=TOB2CleDoc(TOBPiece);
  if TOBREAJ_O.detail.count > 0 then
  begin
    // TODO SUPPR. des anciennes conso.
    ExecuteSQL('DELETE FROM CONSOMMATIONS WHERE BCO_NUMMOUV IN (SELECT BLR_NUMMOUV FROM LIGREAANAL WHERE '+WherePiece(Cledoc,ttdReajAnal,false)+')');
    TOBREAJ_O.DeleteDB(false);
  end;
  for II := 0 to TOBREAJ_O.detail.count -1 do
  begin
    TOBR := TOBREAJ_O.detail[II];
    if TOBR.GetString('BLR_REFANALYTIQUE')<> '' then
    begin
      DecodeRefAna(TOBR.GetString('BLR_REFANALYTIQUE'),RANA);
      DeleteEcrAna (RANA);
    end;
  end;
end;

procedure TOF_BTREAJUSTEANA.AjouteConsommation(TOBD,TOBLigne: TOB);
var
    Part : string;
    Part0 : string;
    Part1 : string;
    Part2 : string;
    Part3 : string;
    Part4 : string;
    DateMouv : TDateTime;
    Annee,Mois,jour : Word;
    Semaine : Integer;
    TypeArticle,Affaire : string;
    coefFG,CoefMarg : Double;
    TOBC : TOB;
    SQL,Nature : string;
    QQ : TQuery;
    NumMouv   : Double;
    TheRetour : TGncERROR;
begin
  TOBD.SetInteger('BLR_NUMMOUV',0);
  Affaire := TOBLigne.GetValue('GL_AFFAIRE');
  if  Affaire = '' then exit;
  //
  Part  := Affaire;
  Part0 := '';
  Part1 := '';
  Part2 := '';
  Part3 := '';
  Part4 := TobLigne.GetValue('GL_AVENANT');

  DateMouv := DateMvt;

  // Découpage du code Affaire
{$IFDEF BTP}
  BTPCodeAffaireDecoupe(Part,Part0,Part1,Part2,Part3, Part4, TaModif,false);
{$ELSE}
  CodeAffaireDecoupe(Part,Part0,Part1,Part2,Part3, Part4, TaModif,false);
{$ENDIF}

  // Découpage de la date
  DecodeDate(DateMouv, Annee, Mois, Jour);

  // Recherche du Numéro de semaine
  Semaine := NumSemaine(DateMouv);
  TypeArticle := TobLigne.GetValue('GL_TYPEARTICLE');
  GetCoefPoc (Affaire,coefFG,CoefMarg);
  if coeffg = 0 then Coeffg := 1;
  if coefmarg = 0 then coefmarg := 1;

  TOBC := TOB.Create ('CONSOMMATIONS',TOBConso,-1);

  TOBC.PutValue('BCO_AFFAIRE', Affaire);
  TOBC.PutValue('BCO_AFFAIRE0', Part0);
  TOBC.PutValue('BCO_AFFAIRE1', Part1);
  TOBC.PutValue('BCO_AFFAIRE2', Part2);
  TOBC.PutValue('BCO_AFFAIRE3', Part3);
  TOBC.PutValue('BCO_PHASETRA', '');
  TOBC.PutValue('BCO_MOIS', Mois);
  TOBC.PutValue('BCO_SEMAINE', Semaine);
  TOBC.PutValue('BCO_DATEMOUV', DateMouv);

  if TypeArticle = 'PRE' then
  Begin
    SQL := 'SELECT GA_ARTICLE,GA_LIBELLE,N1.BNP_TYPERESSOURCE,N1.BNP_LIBELLE FROM ARTICLE ART ' +
               'LEFT JOIN NATUREPREST N1 ON ART.GA_NATUREPRES=N1.BNP_NATUREPRES ' +
               'WHERE GA_TYPEARTICLE="PRE" AND GA_ARTICLE="'+TOBLigne.GetValue('GL_ARTICLE')+'"';
    QQ := OpenSQL(SQL, True);
    TRY
      if not QQ.eof then
      begin
        Nature := QQ.findfield('BNP_TYPERESSOURCE').asString ;
      end else
      begin
        raise Exception.Create('prestation : '+TOBLigne.GetValue('GL_ARTICLE')+' non reconnue');
        exit;
      end;
    FINALLY
      ferme (QQ);
    end;
  end;

  TheRetour := GetNumUniqueConso (NumMouv);
  if TheRetour = gncAbort then
  BEGIN
    raise Exception.Create('Numérotation des consommations');
    Exit;
  END;


//INT,ST,AUT,LOC
  if TypeArticle = 'PRE' then
  begin
    if Nature = 'SAL' then
      TOBC.PutValue('BCO_NATUREMOUV', 'MO')
    else if  Nature = 'ST' then
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
    else if  Nature = 'INT' then
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
    else if  Nature = 'AUT' then
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
    else if  Nature = 'LOC' then
      TOBC.PutValue('BCO_NATUREMOUV', 'EXT')
    else if  Nature = 'MAT' then
      TOBC.PutValue('BCO_NATUREMOUV', 'MAT')
    else if  Nature = 'OUT' then
      TOBC.PutValue('BCO_NATUREMOUV', 'MAT');
  end else if (TypeArticle = 'MAR') or (TypeArticle = 'ARP') then
  Begin
    TOBC.PutValue('BCO_NATUREMOUV', 'FOU');
  end else if TypeArticle = 'FRA' then
  Begin
    TOBC.PutValue('BCO_NATUREMOUV', 'FRS');
  end;

  TOBC.PutValue('BCO_LIBELLE', 'Réajustement Anal. : '+TobLigne.GetValue('GL_LIBELLE'));
  TOBC.PutValue('BCO_TRANSFORME', '-');
  TOBC.PutValue('BCO_NUMMOUV', NumMouv);
  TOBC.PutValue('BCO_INDICE',0);
  if TObLigne.getValue('GL_PRIXPOURQTE')= 0 then TOBLigne.putvalue('GL_PRIXPOURQTE',1);
  TOBC.PutValue('BCO_CODEARTICLE', TobLigne.GetValue('GL_CODEARTICLE'));
  TOBC.PutValue('BCO_ARTICLE', TobLigne.GetValue('GL_ARTICLE'));
  TOBC.PutValue('BCO_QUANTITE', 1);
  TOBC.PutValue('BCO_QUALIFQTEMOUV', TobLigne.GetValue('GL_QUALIFQTEVTE'));
  TOBC.PutValue('BCO_FAMILLENIV1', TobLigne.GetValue('GL_FAMILLENIV1'));
  TOBC.PutValue('BCO_FAMILLENIV2', TOBD.GetString('BLR_FAMILLENIV2'));
  TOBC.PutValue('BCO_FAMILLENIV3', '');

  TOBC.PutValue('BCO_QUANTITE', TOBD.GetDouble('BLR_QTE'));
  TOBC.PutValue('BCO_DPA', TOBD.GetDouble('BLR_DPA'));
  TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*coefFG,V_PGI.okdecP));
  TOBC.PutValue('BCO_PUHT',Arrondi(TOBC.GEtValue('BCO_DPR')*CoefMarg,V_PGI.okdecP));

  TOBC.PutValue('BCO_FACTURABLE', 'N');
  //
  TOBC.PutValue('BCO_VALIDE', '-');
  TOBC.PutValue('BCO_NATUREPIECEG', 'LBT');
  TOBC.PutValue('BCO_SOUCHE', TobLigne.GetValue('GL_SOUCHE'));
  TOBC.PutValue('BCO_NUMERO', TobLigne.GetValue('GL_NUMERO'));
  TOBC.PutValue('BCO_INDICEG', TobLigne.GetValue('GL_INDICEG'));
  TOBC.PutValue('BCO_NUMORDRE', TobLigne.GetValue('GL_NUMORDRE'));
  CalculeLaLigne(TOBC);
  TOBC.SetAllModifie(True);

  TOBD.Setdouble('BLR_NUMMOUV',NumMouv);

end;

procedure TOF_BTREAJUSTEANA.GSGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var TOBL : TOB;
begin
  if ACol < GS.FixedCols then Exit;
  if Arow < GS.Fixedrows then Exit;
  TOBL := GetTOBLigne(TOBsaisie, ARow);
  if TOBL = nil then Exit;
  if TOBL.detail.Count > 0 then  Canvas.Font.Style := Canvas.Font.Style + [fsbold];
end;

function TOF_BTREAJUSTEANA.ExistDatas: boolean;
var II : Integer;
begin
  result := false;
  for II := 0 to TOBSaisie.detail.count - 1 do
  begin
    if TOBSaisie.detail[II].detaiL.count > 0 then
    begin
      result := true;
      exit;
    end;
  end;
end;

procedure TOF_BTREAJUSTEANA.DecodeRefAna(Chaine: string; var RefAna: TRefAna);
var LocStr : string;
begin
  LocStr := Chaine;
	FillChar(RefAna,SizeOf(RefAna),#0);
  //
  RefAna.Entity := StrToInt(ReadTokenSt(LocStr));
  RefAna.Journal := READTOKENST(LocStr);
  RefAna.Exercice := READTOKENST(LocStr);
  RefAna.Axe := READTOKENST(LocStr);
  RefAna.Piece  := StrToInt(READTOKENST(LocStr));
  RefAna.Ligne  := StrToInt(READTOKENST(LocStr));
end;

function TOF_BTREAJUSTEANA.EncodeRefAna(TOBTT: TOB): string;
begin
  Result := InttoStr(TOBTT.GetInteger('Y_ENTITY'))+';'+
            TOBTT.GetString('Y_JOURNAL') + ';' +
            TOBTT.GetString('Y_EXERCICE') + ';' +
            TOBTT.GetString('Y_AXE') + ';' +
            IntToStr(TOBTT.GetInteger('Y_NUMEROPIECE')) + ';' +
            IntToStr(TOBTT.GetInteger('Y_NUMLIGNE'));
end;

procedure TOF_BTREAJUSTEANA.SoldeLigne(Arow: integer);
var TOBL,TOBD : TOB;
begin

  TOBL := TOBSaisie.detail[GS.row-1]; if TOBL = nil then exit;
  TOBD := TOBL.detail[Arow-1]; if TOBD = nil then exit;
  AjouteMontantLigne (TOBD,'-');
  TOBD.SetDouble('BLR_MONTANTACH',ARRONDI(ECART.Value,DEV.Decimale));
  TOBD.SetDouble('BLR_QTE',ARRONDI(TOBD.GetDouble('BLR_MONTANTACH')/TOBD.GetDouble('BLR_DPA'),V_PGI.okdecQ));
  AjouteMontantLigne (TOBD);
  AfficheLaLigne (GS2,TOBL,ARow,fListeSaisie);
end;

Initialization
  registerclasses ( [ TOF_BTREAJUSTEANA ] ) ;
end.
