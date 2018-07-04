{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 28/03/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSELFAC_MUL ()
            .. : Nouvelle gestion de l'édition de DGD
Mots clefs ... : TOF;BTSELFAC_MUL
*****************************************************************}
Unit BTIMPDGD_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Grids,
     AglInit,
{$IFNDEF EAGLCLIENT}
     db,DBGrids,FE_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,MainEagl,
{$ENDIF}
		 Graphics,
     types,
		 Menus,
     HQry,
     HDB,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     UTOF,
     CalcOlegenericBTP,
     EntGC,
     HTB97,
     SaisUtil,
     AffaireUtil,
     splash,
     BTFactImprTOB,
     BTPREPSELFAC_MUL_TOF,
     UTofAfBaseCodeAffaire,uEntCommun,UtilTOBPiece
      ;

Type
  TResultS = (BTrSOk,BTrsNone,BtrsAnnul);
  TOF_BTIMPDGD_MUL = Class (TOF_AFBASECODEAFFAIRE)
      procedure OnNew                    ; override ;
      procedure OnDelete                 ; override ;
      procedure OnUpdate                 ; override ;
      procedure OnLoad                   ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnDisplay                ; override ;
      procedure OnClose                  ; override ;
      procedure OnCancel                 ; override ;
      procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
    private
    	fTypeFacturation : string;
    	MnzPiece : TmenuItem;
      TOBFacturable,TOBmemoFact,TOBAFACTURER,TOBReference,TheOldFac : TOB;
      Blance : TToolbarButton97;
      Fliste : THDbGrid;
      GS : THgrid;
      SqlBase,ColName : string;
      RefPieceSel : String;
      Avancement : boolean;
      //
      procedure SetEventsGridListe(Etat : boolean);
      procedure GetControls;
      procedure FlisteDblClick (Sender : TObject);
      function Chargefacturation(TOBFacture : TOB) : boolean;
      function WhereRefPieceSel(TOBL : TOB) : string;
      procedure NewFacturation (Naturepieceg,Souche : string; Numero,Indice : integer);
      procedure InitGrilleGS;
      procedure AfficheLagridGS;
      procedure TraiteChampAffaireGrid(Gr: THGrid; Nom: string; Col: integer);
      procedure SetEventsGridGS (etat : boolean);
      procedure FlipThisRow (Sender: TObject; ARow: Integer;var Cancel: Boolean);
      function  Existefac (TOBL : TOB): boolean;
      procedure SetEvents;
      procedure BlanceClick (Sender : Tobject);
      procedure EnregistreSelection;
      procedure AddSelection(TOBP: TOB; NumOrdre: integer);
      procedure Majmemorisation;
      procedure SupprimeMemo;
      procedure BchercheClick;
      procedure beforeSelection;
      function FindFirstSelected: TOB;
      function IsPieceSelectionnable(Arow: integer): boolean;
      function ExisteFacturationOld (var Affaire,Naturepieceg,Souche : string; var Numero,Indice : integer) : TResultS;
      procedure ChargeOldfacturation;
      function ExistDevis(cledoc: r_cledoc): boolean; overload;
      function ExistDevis(TOBL : TOB): boolean; overload;
    	function IsPresentDansFactureOld(var TheRefDevis : string;cledoc: r_cledoc): boolean;
      procedure VoirPiece (Sender : TObject);
      procedure GSPostDrawCell (ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    	procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
    	procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
      procedure  EditeDGD ;
      procedure  GetLastInvoice (var Cledoc : r_cledoc);
    	function BeforeImprimeDGD(cledoc: r_cledoc; Affaire: string): boolean;
      procedure ControleChamp(Champ, Valeur: String);

    //
  end ;


Implementation
uses factcomm,FactTOB,FactTvaMilliem,FactCOmmBTP,facture,UtilPgi,UtilArticle,ParamSoc,UTOFBTClotureDev;

procedure TOF_BTIMPDGD_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPDGD_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPDGD_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPDGD_MUL.OnLoad ;
begin
  Inherited ;
  GS.visible := false;
end ;

procedure TOF_BTIMPDGD_MUL.OnArgument (S : String ) ;
var Critere  : string;
    ValMul   : string;
    ChampMul : string;
    x        : integer;
begin
  Inherited ;

  Repeat
    Critere:=uppercase(ReadTokenSt(stArgument)) ;
    valMul := '';
    if Critere<>'' then
    begin
      x:=pos('=',Critere);
      if x<>0 then
      begin
        ChampMul:=copy(Critere,1,x-1);
        ValMul:=copy(Critere,x+1,length(Critere));
      end
      else
        ChampMul := Critere;
      ControleChamp(ChampMul, ValMul);
    end;
  until Critere='';
  //
  TheOldFac := TOB.Create ('LA DERNIERE FACTURE EMISE',nil,-1);
  SqlBase := 'SELECT GP_AFFAIRE,GP_NATUREPIECEG,GP_DATEPIECE,GP_NUMPIECE,GP_PRENUMERO,GP_NUMERO,GP_SOUCHE,GP_INDICEG,GP_TIERS,'+
             'GP_ETABLISSEMENT,GP_REFINTERNE,GP_REPRESENTANT,GP_APPORTEUR,GP_DEVISE,GP_TOTALTTC,GP_TOTALHT,GP_ETATVISA,GP_VIVANTE,'+
             'GP_EDITEE,GP_AFFAIRE1,GP_AFFAIRE2,GP_AFFAIRE3 ,GP_AVENANT,GP_AFFAIREDEVIS,AFF_AFFAIRE1,AFF_AFFAIRE2,'+
             'AFF_AFFAIRE3 ,AFF_AFFAIRE0,AFF_AVENANT,AFF_AFFAIRE,AFF_ETATAFFAIRE, AFF_GENERAUTO,AFF_ADMINISTRATIF, AFF_MODELE,AFF_OKSIZERO '+
             'FROM PIECE LEFT OUTER JOIN AFFAIRE on AFF_AFFAIRE=GP_AFFAIREDEVIS ';
  //
  TOBFacturable := TOB.create ('LA LISTE DES DEVIS',nil,-1);
  TOBmemoFact := TOB.Create ('LES MEMORISATIONS',nil,-1);
  TOBAFACTURER := TOB.Create ('LES PIECES',nil,-1);

  GetControls;
  SetEventsGridListe(True);  // sur Fliste
  SetEvents;
  InitGrilleGS;

  GS.visible := false;

end ;

Procedure TOF_BTIMPDGD_MUL.ControleChamp(Champ : String;Valeur : String);
Begin

  If Champ = 'STATUT' then
  Begin
    if Valeur = 'APP' then
    Begin
      //
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'W');
      //
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="W"');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
      SetControlText('TAFF_AFFAIRE', 'Appel');
    end
    Else if valeur = 'INT' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'I');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="I"');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
      SetControlText('TAFF_AFFAIRE', 'Contrat');
    end
    Else if valeur = 'AFF' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'A')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'A');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A", "")');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
      SetControlText('TAFF_AFFAIRE', 'Chantier');
    end
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControltext('XX_WHERE',' AND AFF_AFFAIRE0 IN ("W", "A")')
      else if assigned(GetControl('AFFAIRE0')) then SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("W","A")');
    end
    Else if valeur = 'PRO' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'P');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="P"');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
      SetControlText('TAFF_AFFAIRE', 'Appel d''offre');
    end
    else
    Begin
      SetControltext('XX_WHERE','');
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', '')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', '');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlText('TAFF_AFFAIRE', 'Affaire');
    end
  end;

end;

procedure TOF_BTIMPDGD_MUL.OnClose ;
begin
  TheOldFac.free;
  TOBFacturable.free;
  TOBmemoFact.free;
  TOBAFACTURER.free;
  Inherited ;
end ;

procedure TOF_BTIMPDGD_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPDGD_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPDGD_MUL.SetEventsGridListe(Etat: boolean);
begin
  if Etat then
  begin
  	Fliste.OnDblClick := FlisteDblClick
  end else
  begin
  	Fliste.OnDblClick := nil;
  end;
end;

procedure TOF_BTIMPDGD_MUL.GetControls;
Var CC : THValComboBox;
begin
  Fliste := THDbGrid(GetCOntrol('Fliste'));
  GS := THgrid(GetCOntrol('GS'));
  Blance := TToolbarButton97 (GetControl('BOUVRIR'));
  GS.visible := false;
	MnzPiece := TmenuItem (GetControl('mnZPiece'));
  //
  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;
  //
  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

end;

procedure TOF_BTIMPDGD_MUL.FlisteDblClick(Sender: TObject);
var Affaire,Naturepieceg,Souche,TypeFac,Sql : string;
    Numero,IndiceG : integer;
    QQ : TQuery;
    TOBFacture : TOB;
    DevisPrinc : string;
    MessAttente : TFsplashScreen;
    sortie : TResultS;
begin
  //
  TheOldFac.clearDetail;
  TOBFacturable.clearDetail;
  //
  DevisPrinc := '';
  TOBFacture := TOB.Create('LA MEMO FACTURE',nil,-1);
  With TFMul(Ecran) do
  begin
    {$IFDEF EAGLCLIENT}
    Q.TQ.Seek(FListe.Row-1) ;
    {$ENDIF}
    TypeFac := Q.findField('AFF_GENERAUTO').AsString;
    Affaire := Q.FindField('GP_AFFAIRE').asstring;
    NaturepieceG := Q.FindField('GP_NATUREPIECEG').asstring ;
    Souche := Q.FindField('GP_SOUCHE').asstring ;
    Numero := Q.FindField('GP_NUMERO').asInteger ;
    IndiceG := Q.FindField('GP_INDICEG').asInteger ;
  end;
  fTypeFacturation := TypeFac;
  if (Pos(TypeFac,'AVA;DAC;')>0) then
  begin
    Avancement := True;
  end else
  begin
    Avancement := false;
  end;
  RefPieceSel := EncoderefSel (NaturePieceg,Souche,Numero,IndiceG);
  QQ := OpenSql ('SELECT BMF_DEVISPRINC FROM BTMEMOFACTURE WHERE BMF_DEVIS="'+RefPieceSel+'"',true);
  if not QQ.eof then
  begin
    // on change la reference de la piece initiale par le devis initial de la facture
    DevisPrinc  := QQ.findfield('BMF_DEVISPRINC').AsString;
  end;
  ferme(QQ);
  // Phase 1 -- On recup les devis d'une facture existante
  if DevisPrinc <> '' then
  begin
    Sql := 'SELECT * FROM BTMEMOFACTURE WHERE BMF_DEVISPRINC="'+DevisPrinc+'"';
    TOBFacture.LoadDetailDBFromSQL ('BTMEMOFACTURE',Sql,false);
    TRY
      RefPieceSel := DevisPrinc;
      if not Chargefacturation(TOBFacture) then
      begin
      	PgiInfo ('Facturation impossible. La facturation à été cloturé',ecran.Caption);
      	exit;
      end;
    FINALLY
      TOBFacture.free;
    END;
  end else
  begin
    Sortie := ExisteFacturationOld (Affaire,Naturepieceg,Souche,Numero,IndiceG);
    if Sortie = Btrsok then
    begin
        MessAttente := TFsplashScreen.Create (ecran);
        MessAttente.Label1.Caption := 'Récupération de la stucture de la facture en cours ...';
        MessAttente.Animate1.Active := true;
        MessAttente.Show;
        MessAttente.Refresh;
        ChargeOldfacturation;
        MessAttente.free;
    end else if Sortie = BtrsNone then
    begin
      NewFacturation (NaturepieceG,Souche,Numero,IndiceG);
    end else if Sortie = BtrsAnnul then
    begin
      exit;
    end;
  end;
  TheOldFac.clearDetail;
  beforeSelection;
  // Now c'est correct -- On peut afficher la grille de multiselection
  AfficheLagridGS;
  GS.Visible := true;
//  if TOBfacturable.detail.count = 1 then BlanceClick (self);

end;

procedure TOF_BTIMPDGD_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff0 := THEdit(GetControl('AFFAIRE0'));
  Aff := THEdit(GetControl('GP_AFFAIRE'));
  Aff1 := THEdit(GetControl('GP_AFFAIRE1'));
  Aff2 := THEdit(GetControl('GP_AFFAIRE2'));
  Aff3 := THEdit(GetControl('GP_AFFAIRE3'));
  Aff4 := THEdit(GetControl('GP_AVENANT'));

end;

function TOF_BTIMPDGD_MUL.Chargefacturation(TOBFacture: TOB) : boolean;
var Sql : string;
    TOBDevis,TOBL : TOB;
    Indice : integer;
    QQ : TQuery;
begin
	result := true;
  For Indice := 0 to TOBFacture.detail.count -1 do
  begin
    TOBL := TOBFacture.detail[Indice];
    Sql := Sqlbase+' WHERE '+WhereRefPieceSel(TOBL);
    QQ := OPenSql (Sql,True);
    if not QQ.eof then
    begin
      TOBDevis := TOB.Create ('PIECE',nil,-1);
      TOBdevis.AddChampSupValeur ('AFF_GENERAUTO','');
      TOBDevis.selectDb('',QQ);
      TOBDevis.changeparent(TOBFacturable,-1);
      TOBDevis.AddChampSupValeur ('SELECTED','X');
      if indice = 0 then TOBReference := TOBdevis;
    end;
    ferme(QQ);
  end;
end;

function TOF_BTIMPDGD_MUL.WhereRefPieceSel(TOBL: TOB): string;
var cledoc : r_cledoc;
begin
   DecodeRefPiece(TOBL.GetValue('BMF_DEVIS'),Cledoc);
   result := WherePiece (cledoc,ttdpiece,true);
end;

procedure TOF_BTIMPDGD_MUL.NewFacturation(Naturepieceg, Souche: string;Numero, Indice: integer);
var Sql : string;
    Piece : string;
    cledoc : r_cledoc;
    TOBDevis : TOB;
    QQ : Tquery;
begin
   Piece := EncoderefSel (Naturepieceg,Souche,Numero,Indice);
   DecodeRefPiece(Piece,Cledoc);
   Sql := SqlBase + ' WHERE '+WherePiece (cledoc,ttdpiece,true);
   QQ := OpenSql (Sql,true);
   if not QQ.eof then
   begin
      TOBDevis := TOB.Create ('PIECE',nil,-1);
      TOBDevis.selectDb('',QQ);
      TOBDevis.changeparent(TOBFacturable,-1);
      TOBdevis.AddChampSupValeur ('SELECTED','X');
      TOBReference := TOBdevis;
   end;
   Ferme (QQ);
end;

procedure TOF_BTIMPDGD_MUL.InitGrilleGS;
var Col,Dec, FixedWidth, Larg : integer ;
    tal : TAlignment ;
    STitre,St,Ch,stA,FF,Typ,FPerso : string ;
    NomList,FRecordSource,FLien,FSortBy,FFieldList,FLargeur,FAlignement,FParams : string ;
    Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul,OkTri,OkNumCol : boolean ;
    Nom : string;
    lescolonnes : string;
    Indice : integer;
    Ftitre,tt,NC : Hstring;
begin
  FixedWidth:=6;
  GS.ColCount:=60;
  Col:=1; STitre:=' '; ColName:='Fixed';

  With TFMul(Ecran) do
  BEGIN
    NomList:=Q.Liste;
    ChargeHListe(NomList,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC,FPerso,OkTri,OkNumCol);
    GS.Titres.Add(FFieldList) ;
    While Ftitre<> '' do
    BEGIN
      StA:=ReadTokenSt(FAlignement);
      St:=ReadTokenSt(Ftitre);
      Ch:=ReadTokenSt(FFieldList);
      Larg:=ReadTokenI(FLargeur);
      tal:=TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
      if OkVisu then
      BEGIN
        GS.Cells[Col,0]:=St ;
        GS.ColAligns[Col]:=tal;
        GS.ColWidths[Col]:=Larg*GS.Canvas.TextWidth('W') ;
        if OkLib then GS.ColFormats[Col]:='CB=' + Get_Join(Ch);
        Typ:=ChampToType(Ch) ;
        if (Typ='INTEGER') or (Typ='SMALLINT') or (Typ='DOUBLE') then GS.ColFormats[Col]:=FF ;
        STitre:=STitre+';'+St;
        ColName:=ColName+';'+Ch;
        inc (Col);
      END;
    END;
    GS.ColCount:=Col ;
    GS.ColWidths[0]:=FixedWidth;
  END ;
  //
  lescolonnes := ColName;
  Nom := READTOKENST (LesColonnes);
  Indice := 0;
  repeat
    TraiteChampAffaireGrid (GS,Nom,indice);
    Nom := READTOKENST (LesColonnes);
    inc(Indice);
  until nom='';
  GS.Options:=GS.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
  GS.Options:=GS.Options+[GoRowSelect] ;
  GS.MultiSelect := true;
  GS.TwoColors := true;
end;

procedure TOF_BTIMPDGD_MUL.AfficheLagridGS;
var Indice : integer;
begin
  SetEventsGridGS (false);
  //
  GS.ClearSelected;
  GS.VidePile (false);
  GS.rowCount := TOBFacturable.detail.count+1;
  For Indice := 0 to TOBFacturable.detail.count -1 do
  begin
    TOBFacturable.detail[Indice].PutLigneGrid (GS,Indice+1,false,false,ColName);
    GS.Objects [0,Indice+1] := TOBFacturable.detail[Indice];
    if TOBFacturable.detail[Indice].GetValue('SELECTED')='X' then GS.FlipSelection (Indice+1);
  end;
  //
  SetEventsGridGS (True);
end;

Procedure TOF_BTIMPDGD_MUL.TraiteChampAffaireGrid ( Gr : THGrid; Nom : string; Col : integer);
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


procedure TOF_BTIMPDGD_MUL.SetEventsGridGS(etat: boolean);
begin
  if Etat  then
  begin
    GS.OnBeforeFlip := FlipThisRow;
    GS.PostDrawCell := GSPostDrawCell;
    GS.OnRowEnter   := GSRowEnter;
    GS.OnRowExit    := GSRowExit;
  end else
  begin
    GS.OnBeforeFlip := nil;
    GS.PostDrawCell := nil;
    GS.OnRowEnter   := nil;
    GS.OnRowExit    := nil;
  end;
end;

function TOF_BTIMPDGD_MUL.IsPieceSelectionnable(Arow : integer) : boolean;
var TOBL : TOB;
begin
  result := true;
  TOBL := TOB(GS.Objects [0,Arow]);
  if TOBreference <> nil then
  begin
    if TOBreference.getValue('RUPTMILLIEME')<> TOBL.getValue('RUPTMILLIEME') then
    begin
      result := false;
    end;
  end;
  if Existefac (TOBL) then
  BEGIN
    PgiBox('Impossible. Il existe une facturation sur ce devis');
    result := false;
  end;
end;

procedure TOF_BTIMPDGD_MUL.FlipThisRow(Sender: TObject; ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
  if GS.IsSelected (Arow) then
  begin
    // Cas ou l'on veut enlever un devis de la sélection
    // On verifie alors s'il existe des lignes de LIGNEFAC pour ce devis
    TOBL := TOB(GS.Objects [0,Arow]);
    if Existefac (TOBL) then
    BEGIN
      PgiBox('Impossible. Il existe une facturation sur ce devis');
      cancel := true;
    end;
    if TOBL = TOBreference then TOBreference := FindFirstSelected;
  end else
  begin
    If not IsPieceSelectionnable(Arow) then begin cancel := true;exit; end;
    if TOBreference = nil then TOBreference := TOB(GS.Objects [0,Arow]);
  end;
end;

function TOF_BTIMPDGD_MUL.FindFirstSelected : TOB;
var Indice : integer;
begin
  result := nil;
  for Indice := 1 to GS.rowCount do
  begin
    if GS.IsSelected (Indice) then
    begin
      result := TOB(GS.Objects [0,Indice]);
      break;
    end;
  end;
end;

function TOF_BTIMPDGD_MUL.Existefac(TOBL: TOB): boolean;
var req : string;
begin
  result := false;
  req := 'SELECT ##TOP 1## BLF_NUMORDRE FROM LIGNEFAC WHERE BLF_NATUREPIECEG="'+
         TOBL.GetValue('GP_NATUREPIECEG')+'" AND BLF_SOUCHE="'+TOBL.GetValue('GP_SOUCHE')+'" '+
         'AND BLF_NUMERO='+InttoStr(TOBL.GetValue('GP_NUMERO'))+' AND BLF_INDICEG='+InttoStr(TOBL.GetValue('GP_INDICEG'));
  if ExisteSql(Req) then BEGIN result := True; Exit; END;
  req := 'SELECT ##TOP 1## BST_NUMEROSIT FROM BSITUATIONS WHERE BST_SSAFFAIRE="'+ TOBL.GetValue('GP_AFFAIREDEVIS')+'"';
  if ExisteSql(Req) then BEGIN result := True; Exit; END;
end;

procedure TOF_BTIMPDGD_MUL.SetEvents;
begin
  Blance.OnClick := BlanceClick;
  mnZPiece.OnClick := VoirPiece;
end;

procedure TOF_BTIMPDGD_MUL.BlanceClick(Sender: Tobject);
begin
  if GS.visible then
  begin
  	EnregistreSelection;
  end else
  begin
    PgiBox('Vous n''avez pas encore sélectionné de devis');
  end;
end;

procedure TOF_BTIMPDGD_MUL.EnregistreSelection;
var Arow,Numordre : integer;
    TOBP : TOB;
    io: TIoErr;
begin
  Numordre := 0;
  TOBmemoFact.clearDetail;
  TOBAfacturer.clearDetail;
  for Arow := 1 to GS.rowCount do
  begin
    if GS.IsSelected (Arow) then
    begin
      Inc(NumOrdre);
      TOBP := TOB(GS.objects[0,Arow]);
      AddSelection(TOBP,NumOrdre);
    end;
  end;
  io := TRANSACTIONS (Majmemorisation,0);
  if io = OeOk then
  begin
    if TOBAFacturer.detail.count > 0 then
    begin
      EditeDGD;
    end;
    BchercheClick;
  end else
  begin
    PGIError ('ATTENTION : Sélection non enregistrée');
    BchercheClick;
  end;
end;

procedure TOF_BTIMPDGD_MUL.AddSelection(TOBP : TOB; NumOrdre : integer);
var TOBMF : TOB;
    Document : string;
    QQ : TQuery;
    piece : string;
    cledoc : r_cledoc;
begin
  // Constitution de la Memorisation
  TOBMF := TOB.Create ('BTMEMOFACTURE',TOBmemoFact,-1);
  TOBMF.putValue('BMF_AFFAIRE',TOBP.GetValue('GP_AFFAIRE'));
  TOBMF.putValue('BMF_AFFAIRE1',TOBP.GetValue('GP_AFFAIRE1'));
  TOBMF.putValue('BMF_AFFAIRE2',TOBP.GetValue('GP_AFFAIRE2'));
  TOBMF.putValue('BMF_AFFAIRE3',TOBP.GetValue('GP_AFFAIRE3'));
  TOBMF.putValue('BMF_DEVISPRINC',RefPieceSel);
  Document := EncoderefSel (TOBP);
  TOBMF.putValue('BMF_DEVIS',document);
  TOBMF.putValue('BMF_NATUREPIECEG',TOBP.GetValue('GP_NATUREPIECEG'));
  TOBMF.putValue('BMF_SOUCHE',TOBP.GetValue('GP_SOUCHE'));
  TOBMF.putValue('BMF_NUMERO',TOBP.GetValue('GP_NUMERO'));
  TOBMF.putValue('BMF_INDICEG',TOBP.GetValue('GP_INDICEG'));
  TOBMF.putValue('BMF_NUMORDRE',NumOrdre);
  // constitution de la TOB des pices a facturer
  TOBMF := TOB.create ('PIECE',TOBAFACTURER,-1);
  Piece := EncodeRefPiece (TOBP);
  DecodeRefPiece (Piece,Cledoc);
  QQ := OpenSql ('SELECT PIECE.*,AFF_GENERAUTO,AFF_OKSIZERO FROM PIECE LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS WHERE '+WherePiece (Cledoc,ttdPiece,false),true);
  if not QQ.eof then
  begin
    TOBMF.SelectDB ('',QQ);
    TOBAfacturer.AddChampSupValeur ('GP_AFFAIRE',TOBMF.GetValue('GP_AFFAIRE'));
    TOBAfacturer.AddChampSupValeur ('GP_TIERS',TOBMF.GetValue('GP_TIERS'));
    TOBAfacturer.AddChampSupValeur ('LIBELLEAFFAIRE',GetChampsAffaire(TOBMF.GetValue('GP_AFFAIRE'),'AFF_LIBELLE'));
    TOBAfacturer.AddChampSupValeur ('LIBELLETIERS',RecupLibelleTiers(TOBMF.GetValue('GP_TIERS')));
    TOBAfacturer.AddChampSupValeur ('AFF_GENERAUTO',TOBMF.GetValue('AFF_GENERAUTO'));
  end;
  ferme(QQ);
end;


procedure TOF_BTIMPDGD_MUL.Majmemorisation;
begin
  SupprimeMemo ;
  if TOBmemofact.detail.count > 0 then if not TOBmemofact.InsertDB (nil) then V_PGI.IOError := OeUnknown;
end;

procedure TOF_BTIMPDGD_MUL.SupprimeMemo;
begin
  V_PGI.ioerror := OeOk;
  ExecuteSql ('DELETE FROM BTMEMOFACTURE WHERE BMF_DEVISPRINC="'+RefPieceSel+'"');
end;

procedure TOF_BTIMPDGD_MUL.BchercheClick;
begin
  TToolBarButton97(GetCOntrol('BCHERCHE')).Click;
end;

procedure TOF_BTIMPDGD_MUL.beforeSelection;
var Indice : integer;
    TOBL,TOBMilliemes : TOB;
    QQ : TQuery;
    cledoc : r_cledoc;
begin
  TOBMilliemes := TOB.Create ('LES MILLIEMEES',nil,-1);
  for Indice := 0 to TOBFacturable.detail.count -1 do
  begin
    TOBL := TOBfacturable.detail[Indice];
 		TOBL.AddChampSupValeur('AVANCSAISIE',0);
 		TOBL.AddChampSupValeur('AVANCPREC',0);
 		TOBL.AddChampSupValeur('RUPTMILLIEME','');
    //
    TOBMilliemes.ClearDetail;
    Cledoc := TOB2CleDoc (TOBL);
    QQ := OPenSql ('SELECT BPM_CATEGORIETAXE,BPM_FAMILLETAXE,BPM_MILLIEME FROM BTPIECEMILIEME WHERE '+WherePiece (Cledoc,TTdRepartmill,false),True);
    if not QQ.eof then
    begin
     TOBMilliemes.LoadDetailDB ('BTPIECEMILIEME','','',QQ,false);
     TOBL.putValue('RUPTMILLIEME',EncodeRepartTva (TOBMilliemes));
    end;
    Ferme(QQ);
  end;
  TOBMilliemes.free;
end;

function TOF_BTIMPDGD_MUL.IsPresentDansFactureOld (Var TheRefDevis : string; cledoc : r_cledoc) : boolean;
var SQl,ClePiece : String;
		QQ : TQuery;
    TOBDFact : TOB;
begin
	result := false;
	TOBDfact := TOB.create ('LES FACTURES', nil,-1);
  TRY
    ClePiece :=';'+Cledoc.NaturePiece+';'+Cledoc.souche+';'+IntToStr(Cledoc.NumeroPiece)+';';
    Sql := 'SELECT DISTINCT BST_SSAFFAIRE FROM LIGNE '+
    			 'LEFT JOIN BSITUATIONS ON BST_NATUREPIECE=GL_NATUREPIECEG AND '+
           'BST_SOUCHE=GL_SOUCHE AND BST_NUMEROFAC=GL_NUMERO WHERE '+
           'GL_NATUREPIECEG="FBT" AND GL_PIECEPRECEDENTE LIKE "%'+ClePiece+'%"';
    QQ := OpenSql (Sql,true,-1,'',True);
    if not QQ.eof then
    begin
      TOBDfact.LoadDetailDB('UNE LIGNE','','',QQ,false);
      TheRefDevis := TOBDfact.detail[0].getValue('BST_SSAFFAIRE');
      result := true;
    end;
    ferme (QQ);
  FINALLY
  	TOBDFact.free;
  END;
end;

function TOF_BTIMPDGD_MUL.ExisteFacturationOld(var Affaire,Naturepieceg,Souche: string; var Numero, Indice: integer): TResultS;
var QQ : TQuery;
    TheRefDevis : string;
    TOBTMP,TOBSEL : TOB;
    req : string;
    cledoc : r_cledoc;
begin
  result := BtrsNone;
  TOBTMP := TOB.Create ('LES SITUATIONS SUR LAFFAIRE',nil,-1);
  TheRefDevis := '';
  QQ := OpenSql ('SELECT GP_AFFAIREDEVIS FROM PIECE WHERE GP_NATUREPIECEG="'+NaturepIeceg+'" AND '+
                 'GP_SOUCHE="'+SOuche+'" AND GP_NUMERO='+InttoStr(Numero)+' AND GP_INDICEG='+InttOstr(Indice),True,-1,'',true);
  if not QQ.eof then
  begin
    TheRefDevis := QQ.findField('GP_AFFAIREDEVIS').AsString;
  end;
  ferme (QQ);

  if not avancement then
  begin
  	result := BtrsNone;
    exit;
  end;

  if TheRefDevis <> '' then
  begin
    QQ := OpenSQL ('SELECT * FROM BSITUATIONS WHERE '+
                  'BST_SSAFFAIRE="'+TheRefDevis+'" '+
                  'BST_VIVANTE="X" '+
                  'ORDER BY BST_NUMEROFAC',true,-1,'',true);
    if not QQ.eof then
    begin
      TheOldFac.LoadDetailDB ('BSITUATIONS','','',QQ,false);
      result := BtrsOK;
    end;
    ferme (QQ);
  end;
  // cas de figure ou il n'existe pas de facturation précédente avce le devis indiqué comme devis principal
  if result <> Btrsok then
  begin
  	//
    cledoc.NaturePiece := Naturepieceg;
    cledoc.Souche      := Souche;
    cledoc.NumeroPiece := Numero;
    cledoc.Indice := Indice;
    if IsPresentDansFactureOld (TheRefDevis,cledoc) then
    begin
      Naturepieceg := Copy(TheRefDevis,2,3);
      Souche := Copy(TheRefDevis,5,3);
      Numero := Strtoint(Copy(TheRefDevis,8,8));
      Indice := 0;
      result := BtrsOK;
      QQ := OpenSQL ('SELECT * FROM BSITUATIONS WHERE '+
                     'BST_SSAFFAIRE="'+TheRefDevis+'" '+
                     'BST_VIVANTE="X" '+
                     'ORDER BY BST_NUMEROFAC',true,-1,'',true);
      if not QQ.eof then
      begin
        TheOldFac.LoadDetailDB ('BSITUATIONS','','',QQ,false);
        result := BtrsOK;
      end;
      ferme (QQ);
      exit;
    end;
    //
    REq := 'SELECT B1.* FROM BSITUATIONS B1 WHERE B1.BST_AFFAIRE="'+Affaire+'" AND B1.BST_NUMEROSIT=(SELECT MAX(B2.BST_NUMEROSIT) FROM BSITUATIONS B2 WHERE B2.BST_SSAFFAIRE=B1.BST_SSAFFAIRE) ORDER BY B1.BST_DATESIT';
    QQ := OpenSQL (req,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBTMP.loadDetailDb ('SITUATIONS','','',QQ,false);
    end;
    ferme (QQ);
    if TOBTMP.detail.count <> 0 then
    begin
      TOBSEL := ChoixSelectionFacture (TOBTMP);
      if TOBSel <> nil then
      begin
        TheRefDevis := TOBSel.getValue('BST_SSAFFAIRE');
        if TheRefDevis <> '' then
        begin
          QQ := OpenSQL ('SELECT * FROM BSITUATIONS WHERE '+
                         'BST_SSAFFAIRE="'+TheRefDevis+'" '+
                         'BST_VIVANTE="X" '+
                         'ORDER BY BST_NUMEROFAC',true,-1,'',true);
          if not QQ.eof then
          begin
            TheOldFac.LoadDetailDB ('BSITUATIONS','','',QQ,false);
            Naturepieceg := Copy(TheRefDevis,2,3);
            Souche := Copy(TheRefDevis,5,3);
            Numero := Strtoint(Copy(TheRefDevis,8,8));
            Indice := 0;
            result := BtrsOK;
          end;
          ferme (QQ);
        end else
        begin
          TOBSel.free;
        end;
      end else
      begin
        result := BtrsAnnul;
      end;
    end;
  end;
end;

function TOF_BTIMPDGD_MUL.ExistDevis (cledoc : r_cledoc): boolean;
begin
  result :=  (TOBFacturable.FindFirst(['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                                [cledoc.NaturePiece,cledoc.Souche,cledoc.NumeroPiece,Cledoc.Indice],true)<>nil);
end;

procedure TOF_BTIMPDGD_MUL.ChargeOldfacturation;
var TOBDevis,TOBFacture,TheLastFacture,TOBL : TOB;
    indice,IndFac : integer;
    QQ : TQuery;
    cledoc : R_CLEDOC;
    first : boolean;
begin
  first := true;
  TOBfacture := TOB.Create ('LA FACTURE',nil,-1);
  TRY
    for IndFac := 0 to TheOldFac.detail.count -1 do
    begin
      TOBFacture.clearDetail;
      TheLastFacture := TheOldFac.detail[IndFac];
      TOBfacture.LoadDetailDBFromSQL ('LIGNE','SELECT GL_TYPELIGNE,GL_PIECEPRECEDENTE,GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO FROM LIGNE '+
                                      'WHERE GL_NATUREPIECEG="'+TheLastFacture.getValue('BST_NATUREPIECE')+'" AND '+
                                      'GL_SOUCHE="'+TheLastFacture.getValue('BST_SOUCHE')+'" AND '+
                                      'GL_NUMERO='+IntToStr(TheLastFacture.getValue('BST_NUMEROFAC'))+' AND GL_TYPELIGNE="ART"',false);
      if TOBfacture.detail.count <> 0 then
      begin
        for Indice := 0 to TOBFacture.detail.count -1 do
        begin
          TOBL := TOBFacture.detail[Indice];
          if TOBL.getValue('GL_TYPELIGNE')<>'ART' then continue;
          if TOBL.getValue('GL_PIECEPRECEDENTE')='' then continue;
          DecodeRefPiece (TOBL.getValue('GL_PIECEPRECEDENTE'),cledoc);
          //
          if not ExistDevis (Cledoc) then
          begin
            QQ := OpenSql (SqlBase +'WHERE GP_NATUREPIECEG="'+Cledoc.NaturePiece+'" AND '+
                           'GP_SOUCHE="'+Cledoc.Souche+'" AND GP_NUMERO='+InttoStr(cledoc.NumeroPiece)+ ' AND '+
                           'GP_INDICEG='+InttOStr(Cledoc.Indice),true,-1,'',true);
            if not QQ.eof then
            begin
              TOBDevis := TOB.Create ('PIECE',nil,-1);
              TOBDevis.selectDb('',QQ);
              TOBDevis.changeparent(TOBFacturable,-1);
              TOBdevis.AddChampSupValeur ('SELECTED','X');
              if first then
              begin
                TOBReference := TOBdevis;
                first := false;
              end;
            end;
          end;
          ferme(QQ);
        end;
      end;
    end;
  FINALLY
    TOBfacture.free;
  END;
end;

function TOF_BTIMPDGD_MUL.ExistDevis(TOBL: TOB): boolean;
var cledoc: r_cledoc;
    TheRef : string;
begin
  TheRef := EncodeRefPiece (TOBL);
  DecodeRefPiece (TheRef,cledoc);
  result := existDevis (cledoc);
end;


procedure TOF_BTIMPDGD_MUL.VoirPiece(Sender: TObject);
var cledoc : r_cledoc;
		Laction : TactionFiche;
    QQ : TQuery;
begin
  if GS.visible then
  begin
    cledoc.NaturePiece  := TOBFacturable.detail[GS.row-1].getValue('GP_NATUREPIECEG');
    cledoc.Souche  := TOBFacturable.detail[GS.row-1].getValue('GP_SOUCHE');
    cledoc.NumeroPiece  := TOBFacturable.detail[GS.row-1].getValue('GP_NUMERO');
    cledoc.Indice  := TOBFacturable.detail[GS.row-1].getValue('GP_INDICEG');
  end else
  begin
    With TFMul(Ecran) do
    begin
      {$IFDEF EAGLCLIENT}
      Q.TQ.Seek(FListe.Row-1) ;
      {$ENDIF}
      cledoc.NaturePiece := Q.FindField('GP_NATUREPIECEG').asstring ;
      cledoc.Souche := Q.FindField('GP_SOUCHE').asstring ;
      cledoc.NumeroPiece := Q.FindField('GP_NUMERO').asInteger ;
      cledoc.Indice := Q.FindField('GP_INDICEG').asInteger ;
    end;
  end;
  //
  if JaiLeDroitNatureGCModif(CleDoc.NaturePiece) then LAction := taModif else LAction := taConsult;
  //
  SaisiePiece(CleDoc, LAction,'','','',false,false,false,true);
  if GS.visible then
  begin
  	if LAction = taModif then
    begin
      QQ := OpenSql ('SELECT * FROM PIECE WHERE '+WherePiece(Cledoc,ttdpiece,true),true,-1,'',true);
      if not QQ.eof then TOBFacturable.detail[GS.row-1].SelectDB('',QQ);
      ferme (QQ);
      TOBFacturable.detail[GS.row-1].PutLigneGrid (GS,Gs.row,false,false,ColName);
    end;
  end else
  begin
    BchercheClick;
  end;
end;

procedure TOF_BTIMPDGD_MUL.GSPostDrawCell(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var ARect: TRect;
		Triangle: array[0..2] of TPoint;
begin
  if (gdFixed in AState) and (ACol = 0) then
  begin
    Arect := GS.CellRect(Acol, Arow);
    Canvas.Brush.Color := GS.FixedColor;
    Canvas.FillRect(ARect);
    if (ARow = GS.row) then
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

procedure TOF_BTIMPDGD_MUL.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
	GS.InvalidateRow(Ou);
end;

procedure TOF_BTIMPDGD_MUL.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
	GS.InvalidateRow(Ou);
end;

procedure TOF_BTIMPDGD_MUL.EditeDGD;
var Cledoc : r_cledoc;
		Affaire : string;
    IMPRESSIONTOB : TImprPieceViaTOB;
begin
  IMPRESSIONTOB := TImprPieceViaTOB.create (ecran);
  TRY
  	GetLastInvoice (Cledoc); // on recupere le dernier numero de facture
    if cledoc.NumeroPiece <> 0 then
    begin
    	ImpressionTOB.SetDocument (cledoc,True);
      Affaire := ImpressionTOB.TOBPiece.getValue('GP_AFFAIREDEVIS');
      if BeforeImprimeDGD (Cledoc,Affaire) then
      begin
      	ImprimerLaPiece(ImpressionTOB.TOBPiece, ImpressionTOB.TOBTiers, CleDoc, false,True, ImpressionTOB.TobAffaire,IMPRESSIONTOB);
      end;
    end;
  FINALLY
  	IMPRESSIONTOB.free;
  END;
end;

procedure TOF_BTIMPDGD_MUL.GetLastInvoice(var Cledoc: r_cledoc);
var QQ : TQuery;
begin
  FillChar (cledoc,sizeof(cledoc),#0);
  QQ := OpenSql ('SELECT ##TOP 1## * FROM BSITUATIONS WHERE BST_NATUREPIECE="FBT" AND BST_SSAFFAIRE="'+ TOBFacturable.detail[0].GetValue('GP_AFFAIREDEVIS')+'" ORDER BY BST_NUMEROSIT DESC',true,-1,'',true);
  if not QQ.eof then
  begin
  	// récupération de la derniére facture émise
  	Cledoc.NaturePiece := QQ.findField('BST_NATUREPIECE').asString;
  	Cledoc.Souche := QQ.findField('BST_SOUCHE').asString;
  	Cledoc.NumeroPiece := QQ.findField('BST_NUMEROFAC').AsInteger;
  	Cledoc.Indice  := 0;
  end;
  ferme (QQ);
end;

function TOF_BTIMPDGD_MUL.BeforeImprimeDGD(cledoc: r_cledoc; Affaire : string): boolean;
var Ret : String;
begin
  result := true;
  Ret := AGLLanceFiche('BTP', 'BTPARIMPDOC', '', '', 'NATURE=' + CleDoc.NaturePiece +
											';SOUCHE=' + CleDoc.Souche +
  										';NUMERO=' + inttostr(CleDoc.NumeroPiece) +
  										 ';AFFAIRE='+Affaire+';DGD');
  result := (ret='1');
end;

Initialization
  registerclasses ( [ TOF_BTIMPDGD_MUL ] ) ;
end.
