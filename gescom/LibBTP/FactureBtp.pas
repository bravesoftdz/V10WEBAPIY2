unit FactureBtp;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, ExtCtrls, HTB97, StdCtrls, HPanel, UIUtil, Hent1, Menus,
  HSysMenu, Mask, Buttons, HStatus, hmsgbox, UTOF, UtilPGI,
  Hqry, UTOB, HFLabel, Ent1, SaisUtil, LookUp, Math, EntGC,
  HrichOle,
  FactCalc,
  // Modif BTP
  factrg,
  //UMetreArticle,
  //UtilMetres,
  UtilPhases,
  BTPUtil, FactOuvrage,UPlannifChUtil,NomenUtil,LigNomen,UTofBTAnalDev,Splash,
  {$IFDEF EAGLCLIENT}
  maineagl,
  {$ELSE}
  DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main, UserChg, AglIsoflex,
  {$ENDIF}
  AglInit, FactComm, FactCpta,
  {$IFNDEF SANSCOMPTA}
  Devise, Chancel,
  {$ENDIF}
  FactTOB,FactPiece,FactArticle,
  FactUtil,
  ImgList,
  UTOF_BTREACPRDOC,
  UTOF_BTSAISDOCEXT,
  BTFactImprTOB,
  UtilArticle,
  AffaireUtil,
  uEntCommun,
  UtilTOBPiece,
  UDemandePrix,
  UtilsMetresXLS;

  var TheMetredoc : TMetreArt;


function  AfficheDesc (FF : Tform ; TOBpiece : TOB) : boolean;
procedure AjusteSousDetail ( FF : Tform ; TOBPiece, TOBNomenclature, TobOuvrage, TOBTiers, TOBAffaire : TOB);
procedure AppelFraisDetailBtp (FF : TForm ; TobPiece,TOBOuvrage,TOBPorcs : TOB; InclusStInFg : boolean=false);
procedure AfficheTexteLigne(FF : Tform ; TOBPiece : TOB; Arow: integer; Desc: THRichEditOLE);
procedure AfterImprimePieceBtp ( FF : Tform ;TOBPiece, TOBNomenclature, TobOuvrage,
                                 TOBTiers, TOBAffaire,TOBLigneRg : TOB;WithReinitQteInit : boolean=true;
                                 ImpressionViaTOB : TImprPieceViaTOB = nil ) ;
procedure AnalyseDocumentBtp (FF : TForm ; TobArticles,TobPiece,TOBOuvrage : TOB ; Ok_Cotraitance : Boolean=True);
procedure AppliqueChangeDate(IndLigne: Integer ; TOBPiece, TOBOuvrage : TOB);
procedure AppliqueChangeTaxe(IndLigne: Integer; TOBPiece,TOBOuvrage : TOB);
function  AppliqueChangementEnteteBtp (FF : TForm;TheTob,TobPiece,TobOuvrage,TOBPorcs,TOBBases,TOBBasesL : TOB; RegimeTaxe : string ; SavModApplicFg,SavModApplicFC  : boolean; OldModeRegl : string) : boolean;
procedure AvancementAcompte (FF : TForm; TobTiers,TOBPiece,TOBPieceRG,LesAcomptes,TOBAcomptes : TOB);
procedure BeforeImprimePieceBtp (FF : TForm; TOBPiece : TOB; var SAVcol : Integer ; Var SAVrow : integer; var ImprimeOk : boolean; var OkAValider : boolean;ImpressionViaTOB : TImprPieceViaTOB = nil; DejaValidee : Boolean=False);
procedure CalculeAvancementBTP ( FF : TForm ; TobPiece,TOBL : TOB ; Acol,Arow : integer; OldQte : double);
procedure ClickValideAndStayHere (FF : Tform ;TOBPiece, TOBNomenclature, TobOuvrage,TOBTiers, TOBAffaire,TOBLigneRg : TOB);
procedure CpltLigneClickBtp (FF : TForm;indicePv : integer; TOBLO,LaTob,TOBArticles,TOBCatalog,TOBOuvrage,TOBTiers,TOBPiece,TOBrepart,TOBCpta,TOBAnaP : TOB);
procedure CreeDesc ( FF : TForm ; TobPiece : TOB);
procedure DefiniPersonnalisationBtp (FF : TForm ; TOBPiece : TOB; ModifAvanc : boolean= false);
procedure DefiniPieceVivanteBtp (TobPiece : TOB ; var NumMess : integer);
procedure DefiniePiedAchat (FF : TForm; TobPiece : TOB; DEV : Rdevise );
procedure DefiniPopLBtp (FF : TForm; Indice : integer);
procedure DefiniPopYBtp (FF : TForm ; TypeArt : string ; TOBPiece : TOB);
procedure DefiniRowCount ( FF : TForm ;TOBPiece : TOB);
procedure DetruitPieceBtp (TOBPiece_O,TobOuvrage_O,TOBPiece, TOBPieceRG, TOBBasesRG, TOBAcomptes : TOB ;DEV : RDevise);
procedure EnabledPiedBtp (FF : TForm ; TobPiece : TOB);
procedure FindNextArticlesBtp (FF : TFORM ; TOBpiece,TOBArticles,TOBTiers,TobAffaire : TOB);
procedure GestionDetailOuvrage(FF : Tform ; TobPiece : TOB ; Arow: integer; Acol : integer = -1);
procedure InformeMetre (FF : Tform ; TOBL : TOB ; Arect : TRect);
procedure InsertLigneInGrilleBtp (FF : TForm;TobPiece,TobTiers,TOBAffaire : TOB; var Arow : integer ; var Acol : integer);
procedure MBDetailNomenClickBtp (FF : Tform ; IndiceNomen : integer; TOBaffaire, TobPiece,TobOuvrage,TobArticles,TOBBases,TOBBasesL,TOBL,TOBrepart,TobMetres : TOB ; addicted : boolean; DEV : Rdevise; ModeSaisieAch : boolean; CotraitOk : boolean);
procedure PositionneColonnesAchatBtp (FF : TForm);
procedure PositionneColonnesBtp ( FF : TForm ;TOBPiece : TOB);
procedure PositionneDocAchatBtp ( FF : TForm);
function TraitePrixOuvrage(TOBPiece, TOBL, TOBBases, TOBBasesL,TOBOuvrage: TOB; EnHt: boolean; PuFix,PuReference: double; DEV: Rdevise; EnPA: boolean = false; FromReference : boolean=false; AffecteEcartLig : Boolean=true;ForcePaEqualPv : Boolean=false): boolean;
function ZoomOuChoixArtLib(FF : TForm ; TOBPiece,TOBArticles,TOBTiers,TOBAffaire : TOB ;ACol, ARow: integer): Boolean;
function IsOuvrage (TOBL : TOB) : boolean;
procedure AffecteQtelivrable (TOBOld,TOBL,TOBArticles : TOB;QteInit : double; gestionConso : TGestionPhase);
//function ControleQteBTP (TOBL : TOB; RefOrigine : string) : boolean;
function ControleQteBTP (TOBL : TOB; RefOrigine : string; QteSais,QteResteInit : double;var PUA : double; TheAction : TactionFiche) : boolean;
procedure AppliqueDestLivraisonFour (TOBpiece : TOB);
function IsModeAvancement (TobPiece : TOB ) : boolean;
function ControleSaisiePieceBtp (TOBPiece : TOB;AFFAIRE : string;TheAction : TactionFiche; transfopiece,duplic : boolean; Silencieux : Boolean = False;IsBordereau : boolean=false) : boolean;
function ControlQteRecepInAchat (FF : Tform ; TOBL,TOBArt : TOB;Acol,Arow : Integer) : boolean;
function WithParamEdition (Naturepiece : string) : boolean;
function ReajustePrixDevis (TheForm : TForm;TOBTiers,TOBLigneTarif,TOBTarif,TOBBases,TOBBasesL,TOBPiece,TOBArticles,TOBOuvrages,TOBPOrcs: TOB ; DEV : RDevise) : boolean;
function GetQteResteCdeOrigine (TOBL : TOB) : double;
function GetMtResteCdeOrigine (TOBL : TOB) : double;
procedure AppliqueGestionStock (TOBPiece,TOBArticles : TOB);
procedure AffecteTenueStock(TOBL,TOBPiece,TOBArticles : TOB);
// procedure AppliqueCoefFrais (TOBPiece,TOBOuvrage : TOB;Coef : double ; DEV : rdevise; InclusStInFg : boolean=false);
procedure AjouteLesLignesBTP (FF: Tform ;TOBpiece,TOBAffaire,TobTiers,TOBArticles: TOB;TheDoc : T_VisuDoc; GestionConso : TGestionPhase);
procedure TentativeAffectationQtive (TOBL,TOBOuvrage : TOB;EnPa,EnPr,EnHt : boolean;PuSaisi : double ; var SumLoc: double ; DEV:rdevise;ForcePaEqualPv:Boolean=false);
procedure PositionneModeBordereauBTP (FF: Tform; TOBPiece : TOB);
procedure InsereLigneRefPrecedent (TOBPiece : TOB ;OldCledoc : R_Cledoc);
procedure AppliquePrixFromTarif (FF : TForm;TOBPiece,TOBL,TOBBases,TOBBasesL,TOBouvrage : TOB;EnPa,EnHT : boolean;Savprix : double;DEV : RDevise; ARow : integer);
procedure  WarningSiPrevisionnel (FF : Tform;TOBpiece,TOBLienDevCha :TOB);
procedure	WarningFinSiprevisionnel (FF : Tform;TOBpiece,TOBLienDevCha :TOB);
function IsPieceOrigineEtude (TOBL : TOB) : boolean;
procedure ReinitCoefPaPr(TOBL,TOBouvrage : TOB);
procedure ReactiveEtude(TOBPiece : TOB);
function ErreurSiZero (TOBPiece : TOB) : boolean;
procedure RetrouvePieceFraisBtp (var Cledoc2 : r_cledoc;TOBpiece : TOB; var ModeSaisie : TActionFiche;WithCreat : boolean=false);
procedure ControleEtReajustePieceOrigine (TOBPiece : TOB; TOBArticles : TOB);
procedure DefiniePiedstd  (FF : TForm);
procedure PositionneColonnesVenteBtp (FF : TForm);
procedure PositionneDocVenteBtp ( FF : TForm);
function GetCoefFGPiece (TOBPiece : TOB; Dev : Rdevise) : double;
procedure PositionnePiece(TOBP : TOB;Cledoc2 : r_cledoc);
procedure DefinieColonnesAchatBtp ( FF : TForm);
function IsModifiable (TOBPiece : TOB) : boolean;
procedure MemoriseListeSaisieBTP (FF : TForm);
procedure ReinitListeSaisieBTP (FF : TForm);
function PieceUniteautorisee (Nature : string) : boolean;
function ControleMargeBTP (TOBPiece : TOB; AfficheMessage : Boolean=False) : boolean;
function ControleVisaBTP(TOBPiece : TOB) : Boolean;
procedure SetTiersTTC (FF : Tform);
procedure PositionneAncienneFacturation(TOBL : TOB; DEV : Rdevise; Typefacturation : string);
procedure InitMontantfacturation(TOBL : TOB);
procedure InformePVBloque (FF : Tform ; TOBL : TOB ; Arect : TRect;TheChaine : string);
function ControleChantiertermineBTP (TOBPiece : TOB;Affaire : string; TransfoPiece,DuplicPiece : boolean) : boolean;
procedure  PositionneAncienneFacDetOuv(TOBL,TOBOuvrages: TOB; DEV : RDevise ;Typefacturation : string);
procedure AppliqueChangeCoefMargOuvrage (LaTOB,TOBOUvrages : TOB; DEV : Rdevise);

implementation

uses facture,
     FactVariante,
     ParamSoc,
     StockUtil,
     FactGrpBtp,
     FactTarifs,
     factDomaines,
     PiecesRecalculs,
     UtofListeInv, TntGrids,UFonctionsCBP,UspecifPOC,
     ULiquidTva2014;


function ControleChantiertermineBTP (TOBPiece : TOB;Affaire : string; TransfoPiece,DuplicPiece : boolean) : boolean;
var naturepiece : String;
    EtatAffaire : String;
    TypeAffaire : String;
begin

	Naturepiece := TOBPiece.GetValue('GP_NATUREPIECEG');
  TypeAffaire := copy(UpperCase (TOBPiece.GetValue('GP_AFFAIRE')),1,1);
  //
	result      := false; // par defaut

  if (duplicPiece) then exit;

  if (Transfopiece) and (Pos(Naturepiece,'FBT;FBP')>0) then exit; // autorise

  //if (duplicPiece) and (Naturepiece='ABT') then exit; // autorise

  if (Pos(Naturepiece,'DBT;PBT;CBT;CF;BLF;LFR;CFR;ETU;FBC;ABC;FBT;LBT;BFC;FF;AF;AFS;DEF;BFA;BCE;ABT;DAC;ABP')> 0) then
  begin
	  EtatAffaire := GetChampsAffaire(Affaire,'AFF_ETATAFFAIRE');
    if (EtatAffaire='TER') Or
       (EtatAffaire='REF') Or
       (EtatAffaire='CL1') Or
       (EtatAffaire='ANN') {Or
       (EtatAffaire='FAC') Or
       (EtatAffaire='REA') Or
       (EtatAffaire='FIN') Or
       (EtatAffaire='ECO') Or
       (EtatAffaire='ACD') } then
    begin
      if TypeAffaire = 'A' then
       	PgiError ('Saisie impossible, le chantier est Terminé ou Refusé')
      else if TypeAffaire = 'I' then
       	PgiError ('Saisie impossible, le Contrat est Terminé ou Refusé')
      else if TypeAffaire = 'W' then
      begin
        if (EtatAffaire = 'CL1') Or (EtatAffaire = 'ANN') then
         	PgiError ('Saisie impossible, l''Appel est Annulé ou Cloturé');
        {else if (EtatAffaire = 'FAC') then
         	PgiError ('Saisie impossible, l''Appel est Facturé')
        else if (EtatAffaire = 'REA') then
          PgiError ('Saisie impossible, l''Appel est Réalisé')
        else if (EtatAffaire = 'FIN') then
          PgiError ('Saisie impossible, l''Appel est Terminé')
        else if (EtatAffaire = 'ECO') then
          PgiError ('Saisie impossible, l''Appel est en cours')
        else if (EtatAffaire = 'ACD') then
          PgiError ('Saisie impossible, l''Appel est en Attente d''acceptation');}
      end;
      result := true;
    end;
  end;

end;


procedure AppliqueGestionStock (TOBPiece,TOBArticles : TOB);
var Indice : integer;
begin
	For Indice := 0 to TOBPiece.detail.count -1 do
  begin
  	if (TOBPiece.Detail[Indice].GetValue('GL_TYPELIGNE') <>'ART') or
    	 ((TOBPiece.Detail[Indice].GetValue('GL_TYPEARTICLE') <>'MAR') and
        (TOBPiece.Detail[Indice].GetValue('GL_TYPEARTICLE') <>'ARP')) then continue;
		AffecteTenueStock (TOBPiece.Detail[Indice],TOBPiece,TOBArticles);
  end;
end;

procedure AffecteTenueStock(TOBL,TOBPiece,TOBArticles : TOB);
var TOBA : TOB;
		PieceTenue,LigneTenue : boolean;
begin
	Piecetenue := (TOBPiece.getValue('_GERE_EN_STOCK')='X');
	TOBA := TOBARticles.findFirst(['GA_ARTICLE'],[TOBL.GetValue('GL_ARTICLE')],true);
  if TOBA <> nil then Lignetenue := (TOBA.getValue('GA_TENUESTOCK')='X')
  							 else LigneTenue := false;
  if (PieceTenue) and (LigneTenue) then TOBL.PutValue('GL_TENUESTOCK','X')
  																 else TOBL.PutValue('GL_TENUESTOCK','-');
end;

function AfficheDesc (FF : Tform ; TOBpiece : TOB) : boolean;
var TOBL: TOB;
  Lib, RefArt: string;
  FFact : TFFacture;
begin
  FFact := TFFacture (FF);
  with FFAct do
  begin
    result := False;
    TOBL := GetTOBLigne(TOBPiece, GS.Row); // récupération TOB ligne
    if TOBL <> nil then
    begin
      Lib := TOBL.GetValue('GL_BLOCNOTE');
      RefArt := TOBL.GetValue('GL_REFARTSAISIE');
      if (RefArt = '') and (Length(Lib) <> 0) and (Lib <> #$D#$A) then
      begin
        textePosition := GS.row;
        AfficheTexteLigne(FF,TobPiece,GS.row, Descriptif1);
        result := True;
      end;
    end;
  end;
end;

procedure AjusteSousDetail ( FF : Tform ; TOBPiece, TOBNomenclature, TobOuvrage, TOBTiers, TOBAffaire : TOB);
var TypeFacturation: string;
    FFact : TFFacture;
begin
  FFact := TFFacture (FF);
  with FFact do
  begin
    SupLesLibDetail(TOBPiece);
//    LoadLesLibDetail(TOBPiece, TOBNomenclature, TobOuvrage, TOBTiers, TOBAffaire, DEV, SaisieTypeAvanc);
    LoadLesLibDetail(TOBPiece, TOBNomenclature, TobOuvrage, TOBTiers, TOBAffaire, FFact.XTOBTRFPOC, DEV, (SaisieTypeAvanc or ModifAvanc ),TheMetredoc);
    if TOBPiece.Detail.Count >= GS.RowCount - 1 then
    begin
      if (tModeBlocNotes in SaContexte) then
      begin
        GS.RowCount := TOBPiece.Detail.Count + 2;
        if GS.RowCount < NBRowsInit then GS.RowCount := NBRowsInit;
        GS.height := (GS.rowHeights[1] * (GS.Rowcount-LeRgpBesoin.NbLignes)) + (GS.GridLineWidth * (GS.Rowcount-LeRgpBesoin.NbLignes));
      end else
      begin
        TypeFacturation := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
//        if (SaisieTypeAvanc = true) and ((TypeFacturation = 'AVA') or (TypeFacturation = 'DAC')) then
        if ((SaisieTypeAvanc) or (ModifAvanc)) and ((TypeFacturation = 'AVA') or (TypeFacturation = 'DAC')) then
          GS.RowCount := TOBPiece.Detail.Count + 1;
        if GS.RowCount < NBRowsInit then GS.RowCount := NBRowsInit;
      end;
    end;
    NumeroteLignesGC(nil, TobPiece);
  end;
end;

procedure AfficheTexteLigne(FF : Tform ; TOBPiece : TOB; Arow: integer; Desc: THRichEditOLE);
var TOBL: TOB;
  ARect: TRect;
  FFact : TFFacture;
  tt : boolean;
begin
  FFact := TFFacture (FF);
  with FFact do
  begin
    TOBL := GetTOBLigne(TOBPiece, Arow); // récupération TOB ligne
    if (TOBL.GetValue('GL_REFARTSAISIE') = '') and (not IsDebutParagraphe(TOBL)) then
    begin
      if (TOBL <> nil)  then
      begin
				Desc.Clear; Desc.Text := '';
        tt := IsBLobVide (FFact,TOBL,'GL_BLOCNOTE');
        if not tt then StringToRich(Desc, TOBL.GetValue('GL_BLOCNOTE'));
      end;
      // Correction fiche Qualité - 10111
      GS.Enabled := false;
      // --
      ARect := GS.CellRect(GS.Col, Arow);
      Desc.BringToFront;
      Desc.Left := ARect.Left;
      Desc.Top := ARect.Top;
      Desc.top := Desc.top + PGS.Top;
      Desc.height := 94; // par défaut dans le facture
      if GS.RowHeights[ARow] > Desc.Height  then Desc.Height := GS.RowHeights[ARow];
      Desc.Width := GS.ColWidths[SG_LIB];
      Desc.Visible := true;
      //   GS.rowHeights[GS.Row] := Desc.Height ;
      Desc.SetFocus;
    end;
  end;
end;


procedure AfterImprimePieceBtp ( FF : Tform ;TOBPiece, TOBNomenclature, TobOuvrage,
                                 TOBTiers, TOBAffaire,TOBLigneRg : TOB;WithReinitQteInit : boolean=true;ImpressionViaTOB : TImprPieceViaTOB = nil ) ;
var FFact : TFFacture;
    i : integer;
begin
  FFact := TFFacture (FF);
  with FFact do
  begin
    if (TOBLigneRG <> nil) and (TOBLigneRg.detail.count > 0) then TOBLigneRG.cleardetail;
    SupLigneRgUnique;
//    LoadLesLibDetail(TOBPiece, TOBNomenclature, TobOuvrage, TOBTiers, TOBAffaire, DEV, SaisieTypeAvanc);
  	if FFact.modeAudit then FFact.AuditPerf.Debut('Ajout des sous détails');
    LoadLesLibDetail(TOBPiece, TOBNomenclature, TobOuvrage, TOBTiers, TOBAffaire, FFact.XTOBTRFPOC, DEV, (SaisieTypeAvanc or ModifAvanc), TheMetreDoc );
  	if FFact.modeAudit then FFact.AuditPerf.Fin('Ajout des sous détails');
    if TOBPiece.Detail.Count >= GS.RowCount - 1 then
    begin
      DefiniRowCount (FF,Tobpiece);
    end;
  	if FFact.modeAudit then FFact.AuditPerf.Debut('Numérotation du document');
    NumeroteLignesGC(nil, TobPiece);
  	if FFact.modeAudit then FFact.AuditPerf.Fin('Numérotation du document');
(*
    if WithReinitQteInit then
    begin
    	MajQtesAvantSaisie (TOBPiece);
    end;
*)
  	if FFact.modeAudit then FFact.AuditPerf.Debut('Définition TOB en mode origine');
    PositionneTOBOrigine;
  	if FFact.modeAudit then FFact.AuditPerf.Fin('Définition TOB en mode origine');
  	if FFact.modeAudit then FFact.AuditPerf.Debut('Affichage dans la grille');
    for i := 0 to TOBPiece.Detail.Count - 1 do AfficheLaLigne(i + 1);
  	if FFact.modeAudit then FFact.AuditPerf.Fin('Affichage dans la grille');
  	if FFact.modeAudit then FFact.AuditPerf.Debut('Définie pièce comme non modifiée');
    InitPasModif;
  	if FFact.modeAudit then FFact.AuditPerf.Fin('Définie pièce comme non modifiée');
  end;
end;

procedure AnalyseDocumentBtp (FF : TForm ; TobArticles,TobPiece,TOBOuvrage : TOB ; Ok_Cotraitance : Boolean=True);
var FFact       : TFFacture;
    Indice      : Integer;
    TOBLOc      : TOB;
    Parametre   : string;
    Libelle     : string;
    Typecotrait : string;
    TypeAnal    : TsrcAnal;
begin
  //
  FFact := TFFacture (FF);
  with FFact do
  begin
    Indice := GS.Row - 1;
    TOBLoc := TOBPIece.Detail[Indice];
//    if copy(TOBLOC.GetValue('GL_TYPELIGNE'), 1, 2) = 'DP' then

		if IsDebutParagraphe (TOBLOC) then
    begin
      Parametre := Parametre + 'DEBUT=' + intTostr(Indice);
      TypeAnal  := SrcPar;
      Libelle   := FTitrePiece.Caption + ' ' + GP_NUMEROPIECE.Caption + ' Paragraphe : ' + TOBLOC.GetValue('GL_LIBELLE');
    end
		else if IsFinParagraphe (TOBLOC) then
    begin
      Parametre := Parametre + 'FIN=' + intTostr(Indice);
      TypeAnal  := SrcPar;
      Libelle   := FTitrePiece.Caption + ' ' + GP_NUMEROPIECE.Caption + ' Paragraphe : ' + TOBLOC.GetValue('GL_LIBELLE');
    end
    else
    begin
      Parametre := Parametre + 'CURLIG=' + intToStr(Indice);
      TypeAnal := SrcOuv;
      if (TOBLOC.GetValue('GL_TYPEARTICLE') = 'ART') or (TOBLOC.GetValue('GL_TYPEARTICLE') = 'ARP') then
        Libelle := FTitrePiece.Caption + ' ' + GP_NUMEROPIECE.Caption + ' Article : ' + TOBLOC.GetValue('GL_LIBELLE')
      else if (TOBLOC.GetValue('GL_TYPEARTICLE') = 'NOM') or (TOBLOC.GetValue('GL_TYPEARTICLE') = 'OUV') then
        Libelle := FTitrePiece.Caption + ' ' + GP_NUMEROPIECE.Caption + ' Ouvrage : ' + TOBLOC.GetValue('GL_LIBELLE');
    end;
    if Ok_Cotraitance then TypeCotrait := 'COTRAITANT=X';
    EntreeAnalyseDocument(TypeAnal,TOBArticles, TOBPiece, TobOuvrage, ['LIBELLE=' + Libelle, TypeCotrait, Parametre], 3);
  end;

end;

procedure AppliqueChangeTaxe(IndLigne: Integer; TOBPiece,TOBOuvrage : TOB);
var TOBL: TOB;
  CodeTaxe1,CodeTaxe2,CodeTaxe3,CodeTaxe4,CodeTaxe5, LocTaxe1,LocTaxe2,LocTaxe3,LocTaxe4,LocTaxe5: string;
  Niveau: Integer;
  Indice, IndDep, IndFin: integer;
begin
  if IndLigne > 0 then
  begin
    CodeTaxe1 := '';
    CodeTaxe2 := '';
    CodeTaxe3 := '';
    CodeTaxe4 := '';
    CodeTaxe5 := '';
    IndDep := Indligne - 1;
    IndFin := IndLigne - 1;
  end else
  begin
    CodeTaxe1 := TOBPiece.GetValue('FAMILLETAXE1');
    CodeTaxe2 := TOBPiece.GetValue('FAMILLETAXE2');
//    CodeTaxe3 := TOBPiece.GetValue('FAMILLETAXE3');
//    CodeTaxe4 := TOBPiece.GetValue('FAMILLETAXE4');
//    CodeTaxe5 := TOBPiece.GetValue('FAMILLETAXE5');
    IndDep := 0;
    IndFin := TOBPiece.detail.count - 1;
  end;
  for Indice := Inddep to IndFin do
  begin
    TOBL := TOBPiece.detail[Indice];
    if TOBL = nil then exit; // pour l'instant
    if CodeTaxe1 <> '' then LocTaxe1 := CodeTaxe1
    else LocTaxe1 := TOBL.GetValue('GL_FAMILLETAXE1');
    if CodeTaxe2 <> '' then LocTaxe2 := CodeTaxe2
    else LocTaxe2 := TOBL.GetValue('GL_FAMILLETAXE2');
    if CodeTaxe3 <> '' then LocTaxe3 := CodeTaxe3
    else LocTaxe3 := TOBL.GetValue('GL_FAMILLETAXE3');
    if CodeTaxe4 <> '' then LocTaxe4 := CodeTaxe4
    else LocTaxe4 := TOBL.GetValue('GL_FAMILLETAXE4');
    if CodeTaxe5 <> '' then LocTaxe5 := CodeTaxe5
    else LocTaxe5 := TOBL.GetValue('GL_FAMILLETAXE5');
// VARIANTE
(*    if (copy(TOBL.Getvalue('GL_TYPELIGNE'), 1, 2) = 'DP') then *)
    if IsDebutParagraphe(TOBL) then
    begin
      Niveau := TOBL.getValue('GL_NIVEAUIMBRIC');
      TOBL.PutValue('GL_RECALCULER', 'X');
      AppliqueChangeTaxePara(TOBPIece, TOBOuvrage, Indice + 1, LocTaxe1,LocTaxe2,LocTaxe3,LocTaxe4,LocTaxe5, Niveau);
    end
    else if (TOBL.Getvalue('GL_TYPEARTICLE') = 'OUV') or (TOBL.GetValue('GL_TYPEARTICLE') = 'ARP') then
    begin
      TOBL.PutValue('GL_FAMILLETAXE1', LocTaxe1); (* a revoir *)
      TOBL.PutValue('GL_FAMILLETAXE2', LocTaxe2); (* a revoir *)
      TOBL.PutValue('GL_FAMILLETAXE3', LocTaxe3); (* a revoir *)
      TOBL.PutValue('GL_FAMILLETAXE4', LocTaxe4); (* a revoir *)
      TOBL.PutValue('GL_FAMILLETAXE5', LocTaxe5); (* a revoir *)
      TOBL.PutValue('GL_RECALCULER', 'X');
      AppliqueChangeTaxeOuv(TOBPiece, TOBOuvrage, Indice + 1, locTaxe1,locTaxe2,locTaxe3,locTaxe4,locTaxe5);
    end
// VARIANTE
(*
    else if (TOBL.GetValue('GL_TYPEARTICLE') <> 'COM') and
      (copy(TOBL.GetValue('GL_TYPEARTICLE'), 1, 2) <> 'TP') then
*)
    else if (not IsCommentaire(TOBL)) and (not IsFinParagraphe(TOBL)) then
    begin
      TOBL.PutValue('GL_RECALCULER', 'X');
      TOBL.Putvalue('GL_FAMILLETAXE1', LocTaxe1);
      TOBL.Putvalue('GL_FAMILLETAXE2', LocTaxe2);
      TOBL.Putvalue('GL_FAMILLETAXE3', LocTaxe3);
      TOBL.Putvalue('GL_FAMILLETAXE4', LocTaxe4);
      TOBL.Putvalue('GL_FAMILLETAXE5', LocTaxe5);
    end;
  end;
  TOBPiece.PutValue('GP_RECALCULER', 'X');
end;

procedure PositioneFinParagraphe (TOBPiece : TOB; Ind,level : integer; TheCurDate : TdateTime);
var Indice : integer;
		TOBL : TOB;
begin
	for indice := ind to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    if IsFinParagraphe (TOBL,level) then
    begin
    	TOBL.PutValue('GL_DATELIVRAISON',TheCurDate);
      break;
    end;
  end;
end;

procedure VerifParagraphePrecedent (TOBPiece : TOB ;Ind,Niveau : integer; TheCurDate : TdateTime);
var indice,level : integer;
		TOBL : TOB;
begin
  for Indice := ind downto 0 do
  begin
  	TOBL := TOBPiece.detail[Indice];
    if IsFinParagraphe (TOBL,1) then break;
    if IsDebutParagraphe (TOBL) then
    begin
    	level := TOBL.getValue('GL_NIVEAUIMBRIC');
    	if level < Niveau then
      begin
      	if TOBL.getValue('GL_DATELIVRAISON') > TheCurDate then
        begin
        	TOBL.PutValue('GL_DATELIVRAISON',TheCurDate);
          PositioneFinParagraphe (TOBPiece,Indice,level,TheCurDate);
        end;
        if Level = 1 then break;
      end;
    end;
  end;
end;


procedure AppliqueChangeDate(IndLigne: Integer ; TOBPiece, TOBOuvrage : TOB);
var TOBL: TOB;
  TheCurDate: TDateTime;
  Niveau: Integer;
  Indice, IndDep, IndFin: integer;
begin
  if IndLigne > 0 then
  begin
    TheCurDate := TobPiece.Detail[IndLigne - 1].getValue('GL_DATELIVRAISON');
    IndDep := Indligne - 1;
    IndFin := IndLigne - 1;
  end else
  begin
    TheCurDate := TobPiece.getValue('GP_DATELIVRAISON');;
    IndDep := 0;
    IndFin := TOBPiece.detail.count - 1;
  end;
  for Indice := Inddep to IndFin do
  begin
    TOBL := TOBPiece.detail[Indice];
    if TOBL = nil then exit; // pour l'instant
//    if (copy(TOBL.Getvalue('GL_TYPELIGNE'), 1, 2) = 'DP') then
		if IsDebutParagraphe (TOBL) then
    begin
      Niveau := TOBL.getValue('GL_NIVEAUIMBRIC');
      if Niveau > 1 then
      begin
      	VerifParagraphePrecedent (TOBPiece,Indice,Niveau,TheCurDate);
      end;
      AppliqueChangeDateLivPara(TOBPIece, TOBOuvrage, Indice + 1, TheCurDate, Niveau);
    end
//    else if (TOBL.GetValue('GL_TYPEARTICLE') <> 'COM') and
    else if (not IsCommentaire (TOBL)) and
      (TOBL.GetValue('GL_TYPELIGNE') <> 'RL') and
//      (copy(TOBL.GetValue('GL_TYPEARTICLE'), 1, 2) <> 'TP') then
      (not IsFinParagraphe (TOBL)) then
    begin
      TOBL.Putvalue('GL_DATELIVRAISON', TheCurDate);
    end;
  end;
end;


procedure AppliqueChangeFraisOuv(TOBPiece,TOBL, TOBOuvrage: TOB; NewgestionFrais,NewgestionFC,NewgestionFA: string);
var TOBOUV,TOBO : TOB;
  	indice,IndiceNomen: Integer;
    ApplicationFrais : string;
begin
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if IsPrestationSt(TOBL) then
  begin
    if (NewGestionFrais='-') and (TOBPiece.getValue('GP_APPLICFGST')='X') then ApplicationFrais:= '-'
                                                                          else ApplicationFrais:= 'X';
    TOBL.Putvalue('GLC_NONAPPLICFRAIS', ApplicationFrais);
    if (NewGestionFC='-') and (TOBPiece.getValue('GP_APPLICFCST')='X') then ApplicationFrais:= '-'
                                                                       else ApplicationFrais:= 'X';
    TOBL.Putvalue('GLC_NONAPPLICFC', ApplicationFrais);
  end else
  begin
    TOBL.Putvalue('GLC_NONAPPLICFRAIS', NewgestionFrais);
    TOBL.Putvalue('GLC_NONAPPLICFC', NewgestionFC);
  end;
  TOBL.Putvalue('GLC_NONAPPLICFG', NewgestionFA);
//  if newGestionFrais = 'X' then TOBL.PutValue('GL_DPR',TOBL.GetValue('GL_DPA'));
  if IndiceNomen=0 then exit;
	TOBOuv := TOBOuvrage.detail[IndiceNomen-1];
  for Indice := 0 to TOBOuv.detail.count - 1 do
  begin
    TOBO := TOBOUV.detail[Indice];
    AppliqueChangeFraisDetOuv (TOBpiece,TOBO,NewGestionfrais,NewgestionFC,NewgestionFA);
  end;
end;

procedure AppliqueChangeFraisPara(TOBPIece, TOBOuvrage: TOB; Ligne: integer; NewgestionFrais,NewgestionFC,NewgestionFA: string; Niveau: integer);
var TOBL: TOB;
  indice: Integer;
  Applicfrais : string;
begin
  for Indice := Ligne to TOBPiece.detail.count - 1 do
  begin
    TOBL := TOBPIece.detail[Indice];
    if IsprestationSt(TOBL) then
    begin
      if (NewgestionFrais='-') and (TOBPiece.getValue('GP_APPLICFGST')='X') then ApplicFrais := '-'
                                                                            else ApplicFrais := 'X';
      TOBL.Putvalue('GLC_NONAPPLICFRAIS', ApplicFrais);
      if (NewgestionFC='-') and (TOBPiece.getValue('GP_APPLICFCST')='X') then ApplicFrais := '-'
                                                                         else ApplicFrais := 'X';
      TOBL.Putvalue('GLC_NONAPPLICFC', ApplicFrais);
    end else
    begin
      TOBL.Putvalue('GLC_NONAPPLICFRAIS', NewgestionFrais);
      TOBL.Putvalue('GLC_NONAPPLICFC', NewgestionFC);
    end;
    TOBL.Putvalue('GLC_NONAPPLICFG', NewgestionFA);
//    if newGestionFrais = 'X' then TOBL.PutValue('GL_DPR',TOBL.GetValue('GL_DPA'));
    if IsOUvrage(TOBL) then AppliqueChangeFraisOuv(TOBPiece,TOBL, TOBOuvrage,NewgestionFrais,NewgestionFC,NewgestionFA);
    if (copy(TOBL.getvalue('GL_TYPELIGNE'), 1, 2) = 'TP') and
      (TOBL.getvalue('GL_NIVEAUIMBRIC') = niveau) then break;
  end;
end;

procedure AppliqueChangeFrais (IndLigne: Integer ; TOBPiece, TOBOuvrage : TOB);
var TOBL: TOB;
	NewGestionFrais,NewGestionFA,NewGestionFC : string;
  Niveau: Integer;
begin
	TOBL := TOBPiece.detail[Indligne-1];
  if TOBL = nil then exit;
	NewGestionFrais := TOBL.getValue('GLC_NONAPPLICFRAIS');
	NewGestionFC := TOBL.getValue('GLC_NONAPPLICFC');
	NewGestionFA := TOBL.getValue('GLC_NONAPPLICFG');
  if IsDebutParagraphe (TOBL) then
  begin
    Niveau := TOBL.getValue('GL_NIVEAUIMBRIC');
    AppliqueChangeFraisPara(TOBPIece, TOBOuvrage, IndLigne, NewgestionFrais,NewgestionFC,NewgestionFA, Niveau);
  end else if (IsOuvrage(TOBL)) then
  begin
    AppliqueChangeFraisOuv(TOBPiece,TOBL, TOBOuvrage, NewgestionFrais,NewgestionFC,NewgestionFA);
  end;
end;



//

procedure AppliqueChangeFCOuv(TOBL, TOBOuvrage: TOB; NewgestionFrais: string);

	procedure AppliqueChangeFraisDetOuv (TOBO : TOB; NewGestionfrais : string);
  var indice : integer;
  		TOBD  : TOB;
  begin
  	TOBO.putValue('BLO_NONAPPLICFC',Newgestionfrais);
    if TOBO.Detail.count > 0 then
    begin
    	for Indice := 0 to TOBO.detail.count -1 do
      begin
      	TOBD := TOBO.detail[Indice];
        AppliqueChangeFraisDetOuv (TOBD,NewGestionfrais);
      end;
    end;
  end;

var TOBOUV,TOBO : TOB;
  	indice,IndiceNomen: Integer;
begin
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceNomen=0 then exit;
	TOBOuv := TOBOuvrage.detail[IndiceNomen-1];
  for Indice := 0 to TOBOuv.detail.count - 1 do
  begin
    TOBO := TOBOUV.detail[Indice];
    AppliqueChangeFraisDetOuv (TOBO,NewGestionfrais);
  end;
end;

procedure AppliqueChangeFCPara(TOBPIece, TOBOuvrage: TOB; Ligne: integer; NewgestionFrais: string; Niveau: integer);
var TOBL: TOB;
  indice: Integer;
begin
  for Indice := Ligne to TOBPiece.detail.count - 1 do
  begin
    TOBL := TOBPIece.detail[Indice];
    TOBL.Putvalue('GLC_NONAPPLICFC', NewgestionFrais);
    if IsOUvrage(TOBL) then AppliqueChangeFCOuv(TOBL, TOBOuvrage,NewgestionFrais);
    if (copy(TOBL.getvalue('GL_TYPELIGNE'), 1, 2) = 'TP') and
      (TOBL.getvalue('GL_NIVEAUIMBRIC') = niveau) then break;
  end;
end;
//


procedure AppliqueChangeFC (IndLigne: Integer ; TOBPiece, TOBOuvrage : TOB);
var TOBL: TOB;
	NewGestionFrais : string;
  Niveau : integer;
begin
	TOBL := TOBPiece.detail[Indligne-1];
  if TOBL = nil then exit;
	NewGestionFrais := TOBL.getValue('GLC_NONAPPLICFC');
  if IsDebutParagraphe (TOBL) then
  begin
    Niveau := TOBL.getValue('GL_NIVEAUIMBRIC');
    AppliqueChangeFCPara(TOBPIece, TOBOuvrage, IndLigne, NewgestionFrais, Niveau);
  end else if (IsOuvrage(TOBL)) then
  begin
    AppliqueChangeFCOuv(TOBL, TOBOuvrage, NewgestionFrais);
  end;
end;

procedure AppliqueNewModeGestionFrais (TOBPiece,TOBOuvrage: TOB ;ZoneAGerer : string;  ModeGestion : boolean);
var TOBL : TOB;
		TheModeGestion : string;
    Indice : integer;
begin
	If ModeGestion then TheModeGestion := '-' else TheModeGestion := 'X';
	for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then continue;
		if ZoneAgerer = 'FR' then
    begin
      if IsprestationST (TOBL) then
      begin
    	  TOBL.PutValue('GLC_NONAPPLICFRAIS',TheModegestion);
      end;
    end else if ZoneAgerer = 'FC' then
    begin
      if IsprestationST (TOBL) then
      begin
    	  TOBL.PutValue('GLC_NONAPPLICFC',TheModegestion);
      end;
    end;

    if (TOBL.GetValue('GL_TYPEARTICLE')='OUV') or (TOBL.GetValue('GL_TYPEARTICLE')='ARP') then
    begin
    	AppliqueModeGestionOuvFg (TOBL,TOBOuvrage,ZoneAGerer,ModeGestion);
    end;
  end;
end;

function AppliqueChangementEnteteBtp (FF : TForm;TheTob,TobPiece,TobOuvrage,TOBPorcs,TOBBases,TOBBasesL : TOB; RegimeTaxe : string ; SavModApplicFg,SavModApplicFC  : boolean; OldModeRegl : string) : boolean;
var FFAct : TFFacture;
		change : boolean;
    ApplicFgst,ApplicFcst : boolean;
begin
  FFAct := TFFActure (FF);
  change := false;
  ApplicFcSt := (TOBPiece.getValue('GP_APPLICFCST')='X');
  ApplicFGSt := (TOBPiece.getValue('GP_APPLICFGST')='X');
  With FFact do
  begin
    if TheTOB.getValue('FAMILLETAXE1') <> '' then
    begin
		  change := True;
      AppliqueChangeTaxe(0,TOBpiece,TOBOuvrage);
      TOBPIECE.Putvalue('FAMILLETAXE1', '');
      TOBPIECE.Putvalue('FAMILLETAXE2', '');
      TOBPIECE.Putvalue('FAMILLETAXE3', '');
      TOBPIECE.Putvalue('FAMILLETAXE4', '');
      TOBPIECE.Putvalue('FAMILLETAXE5', '');
      CacheTotalisations(true);
    end;
    if TOBPiece.getValue('GP_REGIMETAXE') <> RegimeTaxe then
    begin
		  change := True;
      PutValueDetail(TobPiece,'GP_REGIMETAXE',TOBPiece.getValue('GP_REGIMETAXE'));
      PutValueDetailOuv(TOBOuvrage, 'BLO_REGIMETAXE', TOBPiece.getValue('GP_REGIMETAXE'));
      AppliqueChangeTaxe(0,TOBpiece,TOBOuvrage);
    end;
    if TobPiece.GetString('GP_MODEREGLE')<> OldModeRegl then
    begin
    	Change := true;
      FFAct.GP_MODEREGLE.Value := TobPiece.GetString('GP_MODEREGLE');
    end;
    if TheTOB.getValue('GP_APPORTEUR') <> '' then
    begin
      PutValueDetail (TOBPiece,'GP_APPORTEUR',TheTOB.GetValue('GP_APPORTEUR'));
    end;

    if ApplicFgSt <> SavModApplicFg then
    begin
      // Application des frais repartis
    	AppliqueNewModeGestionFrais (TOBPiece,TOBOuvrage,'FR',ApplicFgSt);
    	Change := true;
    end;
    if ApplicFcSt <> SavModApplicFc then
    begin
      // application des frais de chantier
    	AppliqueNewModeGestionFrais (TOBPiece,TOBOuvrage,'FC',ApplicFcSt);
    	Change := true;
    end;
  end;
  result := change;
end;

procedure AvancementAcompte (FF : TForm; TobTiers,TOBPiece,TOBPieceRG,LesAcomptes,TOBAcomptes : TOB);
var
  TOBAcc, TOBPieceLoc, TOBAcptes, TOBA: TOB;
  StRegle, stModeTraitement: string;
  Indice: integer;
  FFact : TFFacture;
begin
  FFact := TFFacture (FF);
  if LesAcomptes = nil then exit;
  TOBAcptes := TOB.Create('LACOMPTE', nil, -1);
  TOBPieceLoc := TOB.Create('PIECE', nil, -1);

  for Indice := 0 to LesAcomptes.detail.count - 1 do
  begin
    TOBA := TOB.Create('ACOMPTES', TOBAcptes, -1);
    TOBA.Dupliquer(LesAcomptes.detail[Indice], true, true);
  end;

  TOBPieceLoc.Dupliquer(TOBPiece, true, true);
  TobAcc := Tob.Create('Les acomptes', nil, -1);
  StRegle := '';
  Tob.Create('', TobAcc, -1);
  TobAcc.Detail[0].Dupliquer(TobTiers, False, TRUE, TRUE);
  Tob.Create('', TobAcc.Detail[0], -1);
  TobAcc.Detail[0].Detail[0].Dupliquer(TobPieceLoc, False, TRUE, TRUE);
  TheTob := TobAcc;
  TOBAcptes.ChangeParent(TobAcc.Detail[0].Detail[0], -1);
  // Modif BTP
  TheTob.data := TOBPieceRG;
  // ---
  if TOBPiece.GetValue('GP_AFFAIREDEVIS') = '' then stModeTraitement := ';HDEVIS'
  else if RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS')) = 'AVA' then stModeTraitement := ';AVANCEMENT'
  else if RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS')) = 'DAC' then stModeTraitement := ';MEMOIRE'
  else stModeTraitement := ';DIRECTE';
  stModeTraitement := stModetraitement + ';CHOIXSEL;TRANSFOPIECE;';
  with FFact do
  begin
    AGLLanceFiche('BTP', 'BTACOMPTES', '', '', ActionToString(Action) + StModeTraitement);
    LEsAcomptes.ClearDetail;
    if TOBAcptes.detail.count = 0 then // cas de l'abandon d'application d'acompte
    begin
      for Indice := 0 to TOBAcomptes.detail.count - 1 do
      begin
        TOBA := TOB.Create('ACOMPTES', LesAcomptes, -1);
        TOBA.Dupliquer(TOBAcomptes.detail[Indice], true, true);
        TOBA.putValue('GAC_MONTANT', 0);
        TOBA.putValue('GAC_MONTANTDEV', 0);
      end;
    end else
    begin
      for Indice := 0 to TOBAcptes.detail.count - 1 do
      begin
        TOBA := TOB.Create('ACOMPTES', LesAcomptes, -1);
        TOBA.Dupliquer(TOBAcptes.detail[Indice], true, true);
      end;
    end;
  end;
  TobAcc.Free;
  TOBPieceLoc.free;
end;

function WithParamEdition (Naturepiece : string) : boolean;
begin
  result := false;
  if (Naturepiece = 'DBT') or (NaturePiece = 'ETU') or
  	 (NaturePiece = 'FBP')  or (NaturePiece = 'BAC') or
  	 (NaturePiece = 'FBT')  or (NaturePiece = 'ABT') or
     (NaturePiece = 'ABP') or (NaturePiece = 'B00') or
     (Naturepiece = 'DAC') or (NaturePiece = 'BCE') then result := true;
end;

procedure BeforeImprimePieceBtp (FF : TForm; TOBPiece : TOB; var SAVcol : Integer ; Var SAVrow : integer; var ImprimeOk : boolean; var OkAValider : boolean;ImpressionViaTOB : TImprPieceViaTOB = nil; DejaValidee : Boolean=False);
var FFact : TFFacture;
    Ret : string;
    cledoc : r_cledoc;
begin
	Ret := '';
	OkAValider := false;
  ImprimeOk := false;
  //
  Cledoc.NaturePiece := TOBPiece.getValue('GP_NATUREPIECEG');
  Cledoc.Souche := TOBPiece.getValue('GP_SOUCHE');
  Cledoc.NumeroPiece := TOBPiece.getValue('GP_NUMERO');
  Cledoc.Indice := TOBPiece.getValue('GP_INDICEG');
  //
  if not AlertDemandePrix (cledoc,TttPEdition) then Exit;
  //
  //FV1 : 22/05/2015 - Contrôle pour interdire l'impression sur pièce visé et montant du visa si la nature est paramétrée en ce sens
  //if Not ControleVisaBTP(TobPiece) then
  //Begin
  //  PGIError(TraduireMemoire('Impression impossible pièce en attente de visa !'),TraduireMemoire('Gestion des Visas'));
  //  Exit;
  //end;
  //
  FFAct := TFFacture( FF);
  with FFact do
  begin
    SavCol := GS.col;
    SavRow := GS.row;
    NextPrevControl(FF);
	//Modif pour impression Affaire Contrat
    if (TOBPiece.GetValue('GP_VENTEACHAT') = 'VEN') and (WithParamEdition(Cledoc.NaturePiece))  then
    //if ((TOBPiece.GetValue('GP_VENTEACHAT') = 'VEN') OR (TOBPiece.GetValue('GP_VENTEACHAT') = 'AUT')) and (WithParamEdition(Cledoc.NaturePiece))  then
    begin
      Ret := AGLLanceFiche('BTP', 'BTPARIMPDOC', '', '', 'NATURE=' + CleDoc.NaturePiece + ';SOUCHE=' + CleDoc.Souche + ';NUMERO=' + inttostr(CleDoc.NumeroPiece) + ';AFFAIRE=' +
        										TOBPiece.GetValue('GP_AFFAIREDEVIS'));
      if Ret = '1' then
      begin
        if (Not DejaValidee) and ((ImpressionViaTOB = nil) or (not ImpressionViaTOB.Usable)) then
        begin
          OKAvalider := PieceModifiee;
          Bvalider.enabled := OKAValider;
          ClickValide(True);
          BValider.Enabled := True;
        end;
      end;
    end else
    begin
      Ret := '1';
      if (Not DejaValidee) then
      begin
        OKAvalider := PieceModifiee;
        Bvalider.enabled := OKAValider;
        ClickValide(True);
        BValider.Enabled := True;
      end;
    end;

    if (ImpressionViaTOB <> nil)  then
    begin
      ImpressionTOB.InCreation := (GP_NUMEROPIECE.Caption = HTitres.Mess[10]);
    end;
    
    if (Ret <> '1') then
    begin
      ImprimeOk := False;
    end else ImprimeOk := true;
  end;
end;

procedure CalculeAvancementBTP ( FF : TForm ; TobPiece,TOBL : TOB ; Acol,Arow : integer; OldQte : double);
var  FFact : TFFacture;
    NewQte, Qte, X  : double;
    TypeFac : string;
begin
  FFact := TFFacture(FF);
  with FFact do
  begin
    // Recalculs pour une situation ou une facture directe :
    if (Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC')>0) and
      (TOBPiece.GetValue('GP_AFFAIREDEVIS') <> '') then
    begin
      // recalcul qté cumulée
      TypeFac := RenvoieTypeFact (TOBPiece.getValue('GP_AFFAIREDEVIS'));
      NewQte := Valeur(GS.Cells[ACol, ARow]);
      Qte := TOBL.GetValue('GL_QTESIT');
(*
      Qte := TOBL.GetValue('GL_QTESIT') - OldQte;
      TOBL.PutValue('GL_QTESIT', Qte + NewQte);
*)
      // recalcul % avancement cumulé ou % de facturation
      OldQte := TOBL.GetValue('GL_QTEPREVAVANC');
      if OldQte <> 0 then X := ((Qte + NewQte) / OldQte) * 100
                     else X := 0;
      if TypeFac = 'DIR' then TOBL.PutValue('GL_QTEPREVAVANC',NewQte) else TOBL.PutValue('GL_POURCENTAVANC', X);
    end;
  end;
end;

procedure AppliqueChangeCoefMargOuvrage (LaTOB,TOBOUvrages : TOB; DEV : Rdevise);
var IndiceNomen,Indice : Integer;
    TOBOuvrage : TOB;
    CoefMarg : double;
begin
  CoefMarg := LaTOB.GetValue('GL_COEFMARG');
  IndiceNomen := LaTOB.GetValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBOuvrage := TOBOUvrages.detail[IndiceNOmen-1];
  for Indice := 0 to TOBOUvrage.detail.count -1 do
  begin
    AppliqueCoefMargDetail (LaTOB,TOBOUvrage.detail[Indice],CoefMarg,DEV);
  end;
end;

procedure CpltLigneClickBtp (FF : TForm;indicePv : integer; TOBLO,LaTob,TOBArticles,TOBCatalog,TOBOuvrage,TOBTiers,TOBPiece,TOBrepart,TOBCpta,TOBAnaP : TOB);
var FFact : TFFacture;
    indiceNomen : integer;
    MontantInterm,MontantPr : double;
    EnHt : boolean;
    TobNomen : TOB;
    TypeArticle :string;
    IsRecalcul : boolean;
begin
	if not isExistsArticle (trim(GetParamsoc('SO_BTECARTPMA'))) then
  begin
		PgiBox (TraduireMemoire('L''article d''écart est invalide ou non renseigné#13#10Veuillez le définir'),Traduirememoire('Gestion d''écart'));
  	exit;
  end;
  //
  IsRecalcul := (LaTob.GetValue('GL_FAMILLETAXE1') <> TOBLO.GetValue('GL_FAMILLETAXE1')) or
    	 					(LaTob.GetValue('GL_FAMILLETAXE2') <> TOBLO.GetValue('GL_FAMILLETAXE2')) or
    	 					(LaTob.GetValue('GL_FAMILLETAXE3') <> TOBLO.GetValue('GL_FAMILLETAXE3')) or
    	 					(LaTob.GetValue('GL_FAMILLETAXE4') <> TOBLO.GetValue('GL_FAMILLETAXE4')) or
    	 					(LaTob.GetValue('GL_FAMILLETAXE5') <> TOBLO.GetValue('GL_FAMILLETAXE5')) or
    						(LaTob.GetValue('GLC_NONAPPLICFRAIS') <> TOBLO.GetValue('GLC_NONAPPLICFRAIS')) or
    	 					(LaTob.GetValue('GLC_NONAPPLICFC') <> TOBLO.GetValue('GLC_NONAPPLICFC')) or
    	 					(LaTob.GetValue('GLC_NONAPPLICFG') <> TOBLO.GetValue('GLC_NONAPPLICFG')) or
								(LaTOB.GetValue('GL_COEFMARG')<> TOBLO.GetValue('GL_COEFMARG')) or
    						(TOBLO.GetValue('GL_QTEFACT') <> LaTob.GetValue('GL_QTEFACT')) or
    						(TOBLO.GetValue('GL_DPA') <> LaTob.GetValue('GL_DPA')) or
       					(TOBLO.GetValue('GL_DPR') <> LaTob.GetValue('GL_DPR')) or
       					(TOBLO.GetValeur(IndicePv) <> LaTob.GetValeur(IndicePv));
  //
	TypeArticle := LaTOB.GetValue('GL_TYPEARTICLE');
  EnHT := (TOBPiece.GetValue('GP_FACTUREHT') = 'X');
  FFact := TFFacture (FF);
  with FFAct do
  begin

    if IsRecalcul then
    begin
      deduitLaLigne(TOBLO);
      ZeroLigneMontant(laTob);
    end;

    if (LaTOB.GetValue('GL_DOMAINE') <> TOBLO.GetValue('GL_DOMAINE')) and (Isparagraphe(LaTOB)) then AffecteLeDomaineParag (FF, GS.Row,TOBPiece);

    if (LaTOB.GetValue('GL_DEPOT') <> TOBLO.GetValue('GL_DEPOT')) and (Isparagraphe(LaTOB)) then AffecteLeDepotParag (FF, GS.Row,TOBPiece);

    if (TObRepart.Detail.count= 0 ) and ((LaTob.GetValue('GL_FAMILLETAXE1') <> TOBLO.GetValue('GL_FAMILLETAXE1')) or
    	 (LaTob.GetValue('GL_FAMILLETAXE2') <> TOBLO.GetValue('GL_FAMILLETAXE2')) or
    	 (LaTob.GetValue('GL_FAMILLETAXE3') <> TOBLO.GetValue('GL_FAMILLETAXE3')) or
    	 (LaTob.GetValue('GL_FAMILLETAXE4') <> TOBLO.GetValue('GL_FAMILLETAXE4')) or
    	 (LaTob.GetValue('GL_FAMILLETAXE5') <> TOBLO.GetValue('GL_FAMILLETAXE5'))) then
    begin
      if (Pos(LaTob.GetValue('GL_FAMILLETAXE1'),VH_GC.AutoLiquiTVAST)>0) and (IsAutoLiquidationTva (TOBpiece)) then
      begin
        TOBPiece.SetBoolean('GP_AUTOLIQUID',true);
      end;
      AppliqueChangeTaxe(GS.row,TOBPiece,TOBOuvrage);
    end;

    // Gestion de la modification de la date de livraison
    if (LaTob.GetValue('GL_DATELIVRAISON') <> TOBLO.GetValue('GL_DATELIVRAISON')) then AppliqueChangeDate(GS.row,TOBPiece,TOBOuvrage);
    // Gestion de la modification de la date de livraison
    if (LaTob.GetValue('GLC_NONAPPLICFRAIS') <> TOBLO.GetValue('GLC_NONAPPLICFRAIS')) or
    	 (LaTob.GetValue('GLC_NONAPPLICFC') <> TOBLO.GetValue('GLC_NONAPPLICFC')) or
    	 (LaTob.GetValue('GLC_NONAPPLICFG') <> TOBLO.GetValue('GLC_NONAPPLICFG')) then
    begin
    	AppliqueChangeFrais(GS.row,TOBPiece,TOBOuvrage);
    end;
    //
    if (LaTOB.GetValue('GL_BLOQUETARIF')='-') and (LaTOB.GetValue('GL_COEFMARG')<> TOBLO.GetValue('GL_COEFMARG')) and (IsOuvrage (LaTOB)) then
    begin
      AppliqueChangeCoefMargOuvrage (LaTOB,TOBOUvrage,TFFActure(FF).DEV);
    end;
    //
    if (IsDebutParagraphe(LaTOb)) then LaTOB.PutValue('GL_FAMILLETAXE1', '');
    //
	  // Modification de l'affaire sur la ligne : maj pour compta analytique
    if (TheTob.GetValue('GL_AFFAIRE') <> TOBLO.GetValue('GL_AFFAIRE')) and TFFACTURE(FF).Comptabilise then
    begin
	    MajAnaAffaire(TOBPiece, TOBCatalog, TOBArticles, TOBCpta, TOBTiers, TOBAnaP, TFFACTURE(FF).GS.Row);
   	end;
    //
    if IsRecalcul then
    begin
      if isOuvrage(LaTOB) then TFFacture(FF).RecalculeOuvrage (LaTob);
      AfficheLaLigne(GS.row);
      TobPiece.PutValue('GP_RECALCULER', 'X');
      LaTob.PutValue('GL_RECALCULER','X')
    end;
  end;
end;

procedure CreeDesc ( FF : TForm ; TobPiece : TOB);
var FFact : TFFacture;
		ii : integer;
begin
  FFact := TFFacture (FF);
  with FFact do
  begin
    ii := GS.row;
    if (IsLigneDetail(TOBPiece, nil, GS.row)) then exit;
    ii := GS.row;
    TextePosition := GS.row;
    AfficheTexteLigne(FF, TobPiece,textePosition, Descriptif1);
    ii := GS.row;
  end;
end;

procedure DefiniPersonnalisationBtp (FF : TForm ; TOBPiece : TOB; ModifAvanc : boolean= false);
var FFact : TFFacture;
  TypeFacturation: string;
  Naturepiece : string;
begin
  FFact := TFFacture (FF);
  NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
  with FFAct do
  begin   
    //FV1 : 20/06/2013 - FS#422 - DELABOUDINIERE : Pb si création d'une commande après sélection sur le code chantier.
    if TFFacture(FF).IsVenteAchat = 'VEN' then
    begin
      // Pas d'accès aux codes tiers et affaire si création depuis la fiche affaire
      // ou si saisie de frais détaillés
      if ((Action = TaCreat) and (GP_AFFAIRE.text <> '')) or
        (TOBPiece.GetValue('GP_NATUREPIECEG') = 'FRC') then
      begin
        GP_TIERS.Enabled := False;
        GP_AFFAIRE1.Enabled := False;
        GP_AFFAIRE2.Enabled := False;
        GP_AFFAIRE3.Enabled := False;
        GP_AVENANT.Enabled := False;
        BRechAffaire.Enabled := False;
        BRazAffaire.Enabled := False;
      end;
    end;
    //
    // Accès à la date de livraison en entête uniquement en commandes fournisseurs
    //
    if Pos(Naturepiece,'DEF;CF;CFR;PBT;CBT;CC') <= 0 then
    begin
      GP_DATELIVRAISON.visible := False;
      HGP_DATELIVRAISON.visible := False;
    end;
    //
    if (SaisieTypeAvanc)  then
    begin
      PPied.Visible := false;
//      if (NaturePiece<>'PBT') then
//      begin
//        PPiedAvanc.Visible := true;
//      end;
      end;
    //
    TypeFacturation := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
    //
    if ((SaisieTypeAvanc) or (ModifAvanc)) and
      ((TypeFacturation = 'AVA') or (TypeFacturation = 'DAC')) then
    begin
      if SG_QA <> -1 then
      begin
        if ModifAvanc then // GS.cells[SG_QA, 0] := TraduireMemoire('Qté situation')
                      else GS.cells[SG_QA, 0] := TraduireMemoire('Qté d''avancement');
      end;
      if SG_PCT <> -1 then GS.cells[SG_Pct, 0] := TraduireMemoire('% d''avancement');
      if FinTravaux then TobPiece.SetAllModifie (true);
    end;
    //
  end;
end;

procedure DefiniPieceVivanteBtp (TobPiece : TOB ; var NumMess : integer);
var TypeFac: string;
begin
  TypeFac := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
  if ((Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'FBT;FBP')>0) and (TypeFac = 'AVA')) or
    ((Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'DAC;BAC')>0) and (TypeFac = 'DAC')) then
  begin
  	if GetChampsAffaire (TOBPiece.GetValue('GP_AFFAIREDEVIS'),'AFF_ETATAFFAIRE') = 'TER' then
    begin
      TobPiece.PutValue('GP_VIVANTE', '-');
      exit;
    end;
    (* Nouvelle gestion --> les anciennes situations deviennent modifiables 
    if ((Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'FBT')>0) and (TypeFac = 'AVA')) or
      ((Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'DAC')>0) and (TypeFac = 'DAC')) then
    begin
      if DerniereSituation(TOBPiece,true) = False then TobPiece.PutValue('GP_VIVANTE', '-');
      NumMess := 43;
    end;
    *)
  end;
end;

function GetCoefFGPiece (TOBPiece : TOB; DEV : RDevise) : double;
var TotalAchat,TOTalrevient : double;
begin
  TotalAchat := arrondi(TOBPiece.GetValue('DPA'), Dev.Decimale);
  TotalRevient := arrondi(TOBPiece.GetValue('DPR'), Dev.Decimale);

  if TotalAchat <> 0 then
  begin
    result := TotalRevient / TotalAchat;
  end
  else
  begin
    result := 0;
  end;
end;

procedure DefiniePiedAchat (FF : TForm; TobPiece : TOB; DEV : Rdevise );
var
  CC: THLabel;
  TotalAchat, TotalRevient, TotalVente, CoefFG, TauxFG, CoefMG, TauxMG: Double;
  FFact : TFFacture;
begin
	TotalVente := 0;
  FFact := TFFacture(FF);
  TotalAchat := arrondi(TOBPiece.GetValue('GP_MONTANTPA'), Dev.Decimale);
  TotalRevient := arrondi(TOBPiece.GetValue('GP_MONTANTPR'), Dev.Decimale);

  if (IsPrevisionChantier (TOBPiece)) or (IsCOntreEtude(TOBpiece))  then
  begin
  	ToTalVente := FFact.GetMontantPvDEv;
  end;

  if TotalVente = 0 then
  begin
  	TotalVente := TOBPiece.GetValue('GP_TOTALHTDEV');
  end;

  if TotalAchat <> 0 then
  begin
    CoefFG := TotalRevient / TotalAchat;
    TauxFG := arrondi(((TotalRevient - TotalAchat) / TotalAchat) * 100, 2);
  end
  else
  begin
    CoefFG := 0;
    TauxFG := 0;
  end;
  if GetParamSocSecur('SO_BTGESTIONMARQ', False) then
  begin
    if TotalRevient <> 0 then CoefMG := TotalVente / TotalRevient
                         else CoefMg := 0; 
    if TotalVente <> 0 then TauxMG := arrondi(((TotalVente - TotalRevient) / TotalVente) * 100, 2)
                       else TauxMg := 0;

  end else
  begin
    if TotalRevient <> 0 then
    begin
      CoefMG := TotalVente / TotalRevient;
      TauxMG := arrondi(((TotalVente - TotalRevient) / TotalRevient) * 100, 2);
    end
    else
    begin
      CoefMG := 0;
      TauxMG := 0;
    end;
  end;
  With FFact do
  begin
    CC := THLabel(FindComponent('HGP_TOTALHTDEV'));
    if CC <> nil then CC.Caption := TraduireMemoire('Total Achat');
    CC := THLabel(FindComponent('LGP_TOTALHTDEV'));
    if CC <> nil then CC.Caption := StrF00(TotalAchat, DEV.Decimale);
    CC := THLabel(FindComponent('HGP_TOTALPORTSDEV'));
    if CC <> nil then CC.Caption := TraduireMemoire('Coefficient Frais');
    CC := THLabel(FindComponent('LTOTALPORTSDEV'));
    if CC <> nil then CC.Caption := StrF00(CoefFG, 4);
    CC := THLabel(FindComponent('HGP_TOTALESCDEV'));
    if CC <> nil then CC.Caption := TraduireMemoire('Pourcentage Frais');
    CC := THLabel(FindComponent('LGP_TOTALESCDEV'));
    if CC <> nil then CC.Caption := StrF00(TauxFG, 2);
    CC := THLabel(FindComponent('HGP_TOTALREMISEDEV'));
    if CC <> nil then CC.Caption := TraduireMemoire('Total Revient');
    CC := THLabel(FindComponent('LGP_TOTALREMISEDEV'));
    if CC <> nil then CC.Caption := StrF00(TotalRevient, DEV.Decimale);
    CC := THLabel(FindComponent('HGP_TOTALTAXESDEV'));
    if CC <> nil then CC.Caption := TraduireMemoire('Coefficient Marge');
    CC := THLabel(FindComponent('LTOTALTAXESDEV'));
    if CC <> nil then
    begin
      CC.Caption := StrF00(CoefMG, 4);
      HTotalTaxesDEV.Text := CC.Caption;
    end;
    CC := THLabel(FindComponent('HGP_TOTALTTCDEV'));
    if CC <> nil then CC.Caption := TraduireMemoire('Pourcentage Marge');
    CC := THLabel(FindComponent('LGP_TOTALTTCDEV'));
    if CC <> nil then CC.Caption := StrF00(TauxMG, 2);
    CC := THLabel(FindComponent('HGP_ACOMPTEDEV'));
    if CC <> nil then CC.Caption := TraduireMemoire('Montant Marge');
    CC := THLabel(FindComponent('LGP_ACOMPTEDEV'));
    if CC <> nil then CC.Caption := StrF00(TotalVente - TotalRevient, DEV.Decimale);
    CC := THLabel(FindComponent('HGP_NETAPAYERDEV'));
    if CC <> nil then
    begin
      if (IsPrevisionChantier (TOBPiece)) or (IsCOntreEtude(TOBpiece))  then
        CC.Caption := TraduireMemoire('Total Devis H.T.')
      else
        CC.Caption := TraduireMemoire('Total H.T.');
    end;
    CC := THLabel(FindComponent('LNETAPAYERDEV'));
    if CC <> nil then CC.Caption := StrF00(TotalVente, DEV.Decimale);
    if Pos(TOBPiece.getValue('GP_NATUREPIECEG'),'CBT,LBT;') > 0 then
    begin
      THPanel(FindComponent('Ptotaux1')).visible := false;
    end;

  end;
end;

procedure DefiniePiedstd  (FF : TForm);
var
  CC: THLabel;
  ASS : THNumEdit;
  FFact : TFFacture;
begin
  FFact := TFFacture(FF);
  With FFact do
  begin
    CC := THLabel(FindComponent('HGP_TOTALHTDEV'));
    if CC <> nil then
    begin
    	CC.Caption := TraduireMemoire('H.T.');
      ASS := THNumEdit(FindComponent('GP_TOTALHTDEV')) ;
      AfficheZonePied (ASS);
    end;
    CC := THLabel(FindComponent('HGP_TOTALPORTSDEV'));
    if CC <> nil then
    begin
    	CC.Caption := TraduireMemoire('dont ports/frais');
      ASS := THNumEdit(FindComponent('TOTALPORTSDEV'));
      AfficheZonePied (ASS);
    end;
    CC := THLabel(FindComponent('HGP_TOTALESCDEV'));
    if CC <> nil then
    begin
    	CC.Caption := TraduireMemoire('Escompte');
      ASS := THNumEdit(FindComponent('GP_TOTALESCDEV'));
      AfficheZonePied (ASS);
    end;
    CC := THLabel(FindComponent('HGP_TOTALREMISEDEV'));
    if CC <> nil then
    begin
    	CC.Caption := TraduireMemoire('Remise');
      ASS := THNumEdit(FindComponent('GP_TOTALREMISEDEV'));
      AfficheZonePied (ASS);
    end;
    CC := THLabel(FindComponent('HGP_TOTALTAXESDEV'));
    if CC <> nil then
    begin
    	CC.Caption := TraduireMemoire('Taxes');
      ASS := THNumEdit(FindComponent('GP_TOTALTAXESDEV'));
      AfficheZonePied (ASS);
    end;
    CC := THLabel(FindComponent('HGP_TOTALTTCDEV'));
    if CC <> nil then
    begin
    	CC.Caption := TraduireMemoire('T.T.C.');
      ASS := THNumEdit(FindComponent('GP_TOTALTTCDEV'));
      AfficheZonePied (ASS);
    end;
    CC := THLabel(FindComponent('HGP_ACOMPTEDEV'));
    if CC <> nil then
    begin
    	CC.Caption := TraduireMemoire('Acompte/Règl.');
      ASS := THNumEdit(FindComponent('GP_ACOMPTEDEV'));
      AfficheZonePied (ASS);
    end;
    CC := THLabel(FindComponent('HGP_NETAPAYERDEV'));
    if CC <> nil then
    begin
    	CC.Caption := TraduireMemoire('Net à payer');
      ASS := THNumEdit(FindComponent('NETAPAYERDEV'));
      AfficheZonePied (ASS);
    end;
  end;
end;

procedure DefiniPopLBtp (FF : TForm; Indice : integer);
var FFAct : TFFacture;
begin
  FFact := TFFacture (FF);
  With FFact do
  begin
    if (POPL.Items.Items[Indice].Name = 'TCommentEnt') then
      POPL.Items.Items[Indice].Caption := TraduireMemoire('Accès au texte de début');
    if (POPL.Items.Items[Indice].Name = 'TCommentPied') then
      POPL.Items.Items[Indice].Caption := TraduireMemoire('Accès au texte de fin');
  end;
end;

procedure DefiniRowCount ( FF : TForm ;TOBPiece : TOB);
var FFact : TFFacture;
begin
  FFact := TFFacture(FF);
  With FFact do
  begin
    GS.RowCount := TOBPiece.Detail.Count + 2;
    if (tModeBlocNotes in SaContexte) then
    begin
      if GS.RowCount < NBRowsInit then GS.RowCount := NBRowsInit+LeRgpBesoin.NbLignes;
//      GS.height := (GS.rowHeights[1] * (GS.Rowcount-LeRgpBesoin.NbLignes)) + (GS.GridLineWidth * (GS.Rowcount-LeRgpBesoin.NbLignes));
    end else
    begin
      if ((SaisieTypeAvanc) {or(ModifAvanc)}) then
      begin
        GS.RowCount := TOBPiece.Detail.Count + 1;
        if GS.RowCount < 15 then GS.RowCount := 15;
      end else

        if GS.RowCount < NBRowsInit then GS.RowCount := NBRowsInit+LeRgpBesoin.NbLignes;
      end;
  end;
end;

procedure ReactiveLigne (cledoc : R_Cledoc);
var chaine : string;
begin
//  chaine := 'UPDATE LIGNE SET GL_VIVANTE="X" WHERE GL_NATUREPIECEG="'+cledoc.NaturePiece+'" AND ' +
  chaine := 'UPDATE LIGNE SET GL_VIVANTE="X",GL_QTERESTE=GL_QTEFACT, GL_MTRESTE=GL_MONTANTHTDEV WHERE GL_NATUREPIECEG="'+cledoc.NaturePiece+'" AND ' +
            'GL_SOUCHE="'+ Cledoc.Souche+'" AND ' +
            'GL_NUMERO='+inttostr(cledoc.NumeroPiece)+' AND ' +
            'GL_NUMORDRE='+inttostr(cledoc.numordre);
  ExecuteSql (chaine);
end;

procedure ReactivePrevisionChantier (TOBPiece : TOB);
var indice : integer;
    TOBL ,TOBPiece_O,TOBP: TOB;
    cledoc : R_Cledoc;
begin
  TOBPiece_O := TOB.Create ('LES PIECES ORIGINES',nil,-1);
  TRY
    for Indice := 0 to TOBpiece.detail.count -1 do
    begin
       TOBL := TOBPiece.detail[Indice];
       if TOBL.GetValue('GL_PIECEPRECEDENTE') <> '' then
       begin
         DecodeRefPiece (TOBL.GetValue('GL_PIECEPRECEDENTE'),Cledoc);
         TOBP := TOBPiece_O.FindFirst (['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                                  [Cledoc.Naturepiece,cledoc.souche,Cledoc.numeropiece,Cledoc.indice],true);
         if TOBP = nil then
         begin
           TOBP := TOB.create ('PIECE',TOBPiece_O,-1);
           TOBP.putvalue ('GP_NATUREPIECEG',cledoc.NaturePiece);
           TOBP.putvalue ('GP_SOUCHE',cledoc.Souche);
           TOBP.putvalue ('GP_NUMERO',cledoc.NumeroPiece);
           TOBP.putvalue ('GP_INDICEG',cledoc.Indice);
           TOBP.PutValue('GP_VIVANTE','X');
         end;
         ReactiveLigne (cledoc);
       end;
    end;
    if TOBPiece_O.detail.count > 0 then TOBPiece_O.UpdateDB (true);
  FINALLY
    TOBPiece_O.free;
  END;
end;

procedure DetruitPieceBtp (TOBPiece_O,TobOuvrage_O,TOBPiece, TOBPieceRG, TOBBasesRG, TOBAcomptes : TOB ;DEV : RDevise);
var nb : integer;
begin
  (* if TOBPiece.getValue('GP_NATUREPIECEG') = VH_GC.AFNataffaire then *) BTPSupprimePieceFrais (TOBpiece);
  // on ne supprime que si l'on est sur la piece ayant crée l'affaire du devis
  if (TOBPiece_O.GetValue('GP_AFFAIREDEVIS') <> '') and (copy(TOBPiece_O.GetValue('GP_AFFAIREDEVIS'), 1, 1) = 'Z') and
      ((GetParamSoc('SO_AFNATAFFAIRE') = TOBPiece_O.GetValue('GP_NATUREPIECEG')) or
      	(TOBPiece_O.GetValue('GP_NATUREPIECEG')='BCE') or
      	(TOBPiece_O.GetValue('GP_NATUREPIECEG')='DE')) then
  begin
    // suppression de la sous-affaire associée
    Nb := ExecuteSQL('DELETE FROM AFFAIRE WHERE AFF_AFFAIRE="' + TOBPiece_O.GetValue('GP_AFFAIREDEVIS') + '"');
    if Nb < 0 then
    begin
      V_PGI.IoError := oeUnknown;
      Exit;
    end;
    if not ControleAffaireRef(TOBPiece_O.GetValue('GP_AFFAIRE')) then
    begin
      V_PGI.IoError := oeUnknown;
      Exit;
    end;
  end;
  // suppression dans BTPARDOC
  ExecuteSQL('DELETE FROM BTPARDOC WHERE BPD_NATUREPIECE="' + TOBPiece_O.GetValue('GP_NATUREPIECEG') + '" AND BPD_SOUCHE="' + TOBPiece_O.GetValue('GP_SOUCHE')
    + '" AND BPD_NUMPIECE=' + IntToStr(TOBPiece_O.GetValue('GP_NUMERO')));
  if (GetParamSoc('SO_BTNATCHANTIER') = TOBPIece.GetValue('GP_NATUREPIECEG')) or (TOBPIece.GetValue('GP_NATUREPIECEG')='BCE') then
  begin
    if V_PGI.IoError = oeOk then DeprepareDevis(TOBPiece);
    if V_PGI.IoError = oeOk then DetruitLienDevisChantier(TOBPiece);
  end;
  (*
  // Reactualisation des acomptes dans l'ancienne Piece impossible du fait que l'on peut regrouper
  if TOBPiece.GetValue('GP_NATUREPIECEG')='CBT' then
  begin
    if V_PGI.IoError = oeOk then ReactivePrevisionChantier(TOBPiece);
  end;
  *)
  (* CONSO *)
  if V_PGI.IoError = oeOk then AnnulePhases(TOBPiece);
  (* ---- *)
end;

procedure EnabledPiedBtp (FF : TForm ; TobPiece : TOB);
var FFact : TFFacture;
    TypeFac: string;
    VenteAchat : string;
    ftobPieceTrait : TOB;
begin
  FFact := TFFacture (FF);
  ftobPieceTrait := FFact.ThePieceTrait;
  With FFact do
  begin
    VenteAchat := GetInfoParPiece(NewNature, 'GPP_VENTEACHAT');

    //AnalyseDocHtrait.visible := ((VenteAchat = 'VEN') or (VenteAchat = 'AUT')) and (ftobPieceTrait.detail.count <= 1) ;

    MBModevisu.visible := not OrigineEXCEL;
    MBSGED.visible := False;
    //MBSoldeTousReliquat.visible := False; : modif brl 22/02/2011 on ne sait pas pourquoi le bouton était invisible ...
    if Not SaisieTypeAvanc then
    begin
      BImprimer.visible := True;
      Bminute.visible := ((TOBPiece.GetValue('GP_NATUREPIECEG') = VH_GC.AFNatAffaire) or
                          (TOBPiece.GetValue('GP_NATUREPIECEG') = VH_GC.AFNatProposition) or
                          (TOBPiece.GetValue('GP_NATUREPIECEG') = GetParamSoc('SO_BTNATCHANTIER'))) and (VenteAchat = 'VEN');
      Bminute2.visible := Bminute.visible;
    end;
    if TOBPiece.GetValue('GP_NATUREPIECEG') = VH_GC.AFNatProposition then
      Bchercher.Hint := TraduireMemoire('Recherche et affectation globale');
    BVentil.Visible := False;
    //

    SIMULATIONRENTABILIT1.visible := (VenteAchat = 'VEN');
    SIMULATIONRENTABILIT1.Enabled := not (Action = taConsult);
    //

    Typefac := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
    if ((Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'FBT;FBP')>0) and (Typefac = 'AVA')) or
      ((Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'DAC;BAC')>0) and (Typefac = 'DAC')) then
    begin
      BprixMarche.Visible := False;
      SIMULATIONRENTABILIT1.visible := False;
      MBTarif.Visible := False;
			MBREPARTTVA.Visible := False;
    end;
    if (VenteAchat <> 'VEN') and (BPrixMarche.Visible) then BPrixMarche.Visible := false;
    {$IFNDEF LINE}
    FraisDetail.visible := (Action = taModif);
    FFact.FraisDetail.enabled := (Action = taModif);
    N8.visible := true;
    // Cas du devis facturé
    if (TFFacture(FF).IsDejaFacture) or (TFFacture(FF).IsFromPrepaFac) then
    begin
      SIMULATIONRENTABILIT1.visible := false;
    end;
    if Not ExJaiLeDroitConcept(TConcept(gcSaisModifRepres),False) then FFact.GP_REPRESENTANT.Enabled:=False; //confidentialité , concept acces remise
    if Not ExJaiLeDroitConcept(TConcept(bt513),False) then FFact.GP_REMISEPIED.Enabled:=False; //confidentialité , concept acces remise
    if Not ExJaiLeDroitConcept(TConcept(bt514),False) then FFact.GP_ESCOMPTE.Enabled:=False; //confidentialité , concept acces escompte
    if Not ExJaiLeDroitConcept(TConcept(bt515),False) then FFact.SIMULATIONRENTABILIT1.Enabled:=False; //confidentialité , concept acces simulation
		{$ENDIF}
  end;
end;

procedure DefiniPopYBtp (FF : TForm ; TypeArt : string ; TOBPiece : TOB);
var FFact : TFFacture;
    indice : integer;
begin
  FFact := TFFacture(FF);
  with FFact do
  begin
    MBDetailLot.Visible := False;
    for Indice := 0 to POPY.Items.Count - 1 do
    begin
      if (POPY.Items.Items[Indice].Name = 'MBDetailNomen') then
      begin
        if TypeArt = 'NOM' then
          POPY.Items.Items[Indice].Caption := TraduireMemoire('Détail Nomenclature')
        else
          POPY.Items.Items[Indice].Caption := TraduireMemoire('Détail Ouvrage');
      end;

      if (POPY.Items.Items[Indice].Name = 'CpltEntete') then
        POPY.Items.Items[Indice].Caption := TraduireMemoire('Complément Document');
    end; // fin du FOR
(* modif brl 080604 :
    if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_GESTDETFRAIS') = 'X' then
*)
{$IFNDEF LINE}
    if (isPieceGerableFraisDetail(TOBPiece.getValue('GP_NATUREPIECEG')) and (TOBPiece.getValue('GP_NATUREPIECEG')<>'PBT')) then
      begin
      FraisDetail.Visible := True;
      FraisDetail.Enabled := (Action=TaModif);
      end
    else
      FraisDetail.Visible := False;
{$ENDIF}
  end;

end;

procedure FindNextArticlesBtp (FF : TFORM ; TOBpiece,TOBArticles,TOBTiers,TobAffaire : TOB);
var FFAct : TFFActure;
    ind : integer;
    TOBL : TOB;
begin
  FFact := TFFacture (FF);
  With FFact do
  begin
    TRY
      // Boucle de recherche à partir de la position courante dans la grille
      for Ind := GS.Row - 1 to TOBPiece.detail.count - 1 do
      begin
        TOBL := GetTOBLigne(TOBPiece, Ind + 1);
        // On ne traite que les lignes quantifiées hormis les lignes de sous-détails
        if (TOBL <> nil) and not (IsLigneDetail(nil, TOBL)) and (TOBL.GetValue('GL_QTEFACT') <> 0) then
        begin
          GS.Row := Ind + 1;
          GS.Col := SG_Lib;
          StCellCur := GS.Cells[GS.Col, GS.Row];
          // Si on stoppe la recherche, la fonction ZoomOuChoixArtLib retourne False
          if not (ZoomOuChoixArtLib(FF,TOBPiece,TobArticles,TobTiers,TobAffaire,GS.Col, GS.Row)) then Exit;
        end;
      end;
    FINALLY
    	FFAct.CalculeLaSaisie (-1, -1, True);
    end;

  end;
end;
(*
procedure AffecteCoefDetailOuvrage (TOBOuvrage,TOBL : TOB;DEV : Rdevise;Coef : double; MajPv : boolean; InclusStInFg : boolean);
var Indice : integer;
		CoefMarg : double;
    TOBO : TOB;
    EnHt : boolean;
begin
	EnHT := (TOBL.GetValue('GL_FACTUREHT')='X');
	for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
  	TOBO := TOBOuvrage.detail[Indice];
    if TOBO.detail.Count > 0 then
    begin
    	AffecteCoefDetailOuvrage (TOBO,TOBL,DEV,Coef,MajPv,InclusStInFg);
    end else
    begin
//    	if (TOBO.GetValue('BLO_TYPEARTICLE')='PRE') and (RenvoieTypeRes (TOBO.GEtValue('BLO_ARTICLE'))='ST' ) then continue;
      // Correction de l'anomalie sur les ECART PM
      if (TOBO.GetValue('BLO_DPR')=0) and (TOBO.GetValue('BLO_DPA')=0) and (TOBO.GetValue('BLO_PUHTDEV')<>0) then continue;
      //
    	if TOBO.GetValue('BLO_DPR')=0 then TOBO.putValue('BLO_DPR',TOBO.GetValue('BLO_DPA'));
      if TOBO.GetValue('BLO_DPR')=0 then
      begin
      	if EnHt then TOBO.putValue('BLO_DPR',TOBO.GetValue('BLO_PUHTDEV'))
        				else TOBO.putValue('BLO_DPR',TOBO.GetValue('BLO_PUTTCDEV'));
      end;
      if TOBO.GetValue('BLO_DPR')=0 then
      begin
      	CoefMarg := 1;
      end else
      begin
      	if EnHt then CoefMarg := TOBO.GetValue('BLO_PUHTDEV')/TOBO.getValue('BLO_DPR')
        			  else CoefMarg := TOBO.GetValue('BLO_PUTTCDEV')/TOBO.getValue('BLO_DPR');
      end;
    	if ((TOBO.GEtValue('BNP_TYPERESSOURCE')<>'ST') and (TOBO.GetValue('BLO_NONAPPLICFRAIS')<>'X'))
      		OR
      	 ((TOBO.GEtValue('BNP_TYPERESSOURCE')='ST') and
         (InClusStInFg) and (TOBO.GetValue('BLO_NONAPPLICFRAIS')<>'X'))  then
      begin
      	TOBO.putValue('BLO_DPR',Arrondi(TOBO.GetValue('BLO_DPA')*Coef,V_PGI.okdecP));
      end;
      if MajPV then
      begin
        if EnHt then
        begin
          TOBO.PutValue('BLO_PUHTDEV',Arrondi(TOBO.GetValue('BLO_DPR')*TOBO.GEtValue('__COEFMARG'),V_PGI.OkdecP));
        end else
        begin
          TOBO.PutValue('BLO_PUTTCDEV',Arrondi(TOBO.GetValue('BLO_DPR')*TOBO.GEtValue('__COEFMARG'),V_PGI.OkdecP));
        end;
      end;
    end;
  end;
end;

procedure AppliqueCoefFrais (TOBPiece,TOBOuvrage : TOB;Coef : double ; DEV : rdevise; InclusStInFg : boolean=false);
var Indice : integer;
		TheCoefMarg : double;
    TOBL,TOBO : TOB;
    EnHt : boolean;
    Valeurs : T_Valeurs;
begin
	EnHt := (TOBPiece.GetValue('GP_FACTUREHT')='X');
	for Indice := 0 to TOBpiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    if TOBL.GetVAlue('GL_TYPELIGNE')<>'ART' then continue;
    if ((TOBL.GetValue('GL_TYPEARTICLE')='OUV') or (TOBL.GetValue('GL_TYPEARTICLE')='ARP')) and
    		(TOBL.getValue('GL_INDICENOMEN')>0) then
    begin
      TOBO := TOBOuvrage.detail[TOBL.GetValue('GL_INDICENOMEN') -1];
    	// C'est un ouvrage donc...
      AffecteCoefDetailOuvrage (TOBO,TOBL,DEV,Coef,(TOBL.GEtValue('GL_BLOQUETARIF')='-'),InclusStInFg);
      //
//      gcvoirtob(TOBO,'','BLO_ARTICLE;BLO_LIBELLE;BLO_QTEFACT;BLO_DPA;BLO_DPR;BLO_PUHTDEV');
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBO,1,1,true,DEV,valeurs,EnHt,true,true);
      if TOBL.GEtValue('GL_BLOQUETARIF')='-' Then
      begin
        TOBL.Putvalue('GL_PUHT',DeviseToPivot(valeurs[2],DEV.Taux,DEV.quotite));
        TOBL.Putvalue('GL_PUTTC',DevisetoPivot(valeurs[3],DEV.taux,DEV.quotite));
        TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
        TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
        TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
        TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
      end;
      TOBL.Putvalue('GL_DPA',valeurs[0]);
      TOBL.Putvalue('GL_DPR',valeurs[1]);
      TOBL.Putvalue('GL_PMAP',valeurs[6]);
      TOBL.Putvalue('GL_PMRP',valeurs[7]);
      TOBL.putvalue('TPSUNITAIRE',valeurs[9]);
      TOBL.putvalue('GL_RECALCULER', 'X');
      positionneCoefMarge (TOBL);
      //
    end else if TOBL.GetValue ('GL_TYPELIGNE')='ART' then
    begin
      // correction sur les cart PM
      if (TOBL.GetValue('Gl_DPA')=0) and (TOBL.GetValue('GL_DPR')=0) and (TOBL.GetValue('GL_PUHTDEV')<>0) then continue;
      // --
			positionneCoefMarge (TOBL);
      // Nouvel emplacement
      if ((TOBL.GetValue('BNP_TYPERESSOURCE') <> 'ST') and
      	 (TOBL.GetValue('GL_TYPEARTICLE')<>'OUV') and
      	 (TOBL.GetValue('GL_ARTICLE')<>'') and
         (TOBL.GetValue('GLC_NONAPPLICFRAIS')<>'X'))
         OR
         ((TOBL.GetValue('BNP_TYPERESSOURCE') = 'ST') and (InclusStInFg) and
           (TOBL.GetValue('GLC_NONAPPLICFRAIS')<>'X')) then
     	begin
      	TOBL.putValue('GL_DPR',Arrondi(TOBL.GetValue('GL_DPA')*Coef,V_PGI.okdecV));
     	end;
      //
      if TOBL.GEtValue('GL_BLOQUETARIF')='-' Then
      begin
        if TOBL.GetValue('GL_FACTUREHT')='X' then
        begin
          TOBL.putValue('GL_PUHTDEV',Arrondi(TOBL.GetValue('GL_DPR')* TOBL.GEtValue('__COEFMARG'),V_PGI.okdecP));
          TOBL.putValue('GL_PUHT',DEVISETOEURO (TOBL.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.Quotite)  );
        end else
        begin
          TOBL.putValue('GL_PUTTCDEV',Arrondi(TOBL.GetValue('GL_DPR')*TOBL.GEtValue('__COEFMARG'),V_PGI.okdecP));
          TOBL.putValue('GL_PUTTC',DEVISETOEURO (TOBL.GetValue('GL_PUTTCDEV'),DEV.Taux,DEV.Quotite)  );
        end;
      end;
      TOBL.putvalue('GL_RECALCULER', 'X');
    end;
  end;
  TOBPiece.PutValue('GP_RECALCULER', 'X');
end;
*)
procedure PositionnePiece(TOBP : TOB;Cledoc2 : r_cledoc);
begin
  TOBP.putValue('GP_NATUREPIECEG',Cledoc2.NaturePiece);
  TOBP.PutValue('GP_DATEPIECE',Cledoc2.DatePiece);
  TOBP.PutValue('GP_NUMERO',Cledoc2.NumeroPiece);
  TOBP.PutValue('GP_SOUCHE',Cledoc2.Souche );
  TOBP.PutValue('GP_INDICEG',Cledoc2.Indice );
end;

procedure RetrouvePieceFraisBtp (var Cledoc2 : r_cledoc;TOBpiece : TOB; var ModeSaisie : TActionFiche ;WithCreat : boolean=false );
var TOBP : TOB;
		EnHt ,SaisieContre: boolean;
    Indice : integer;
begin
	if TOBPiece.getValue ('GP_PIECEFRAIS')='' then
  begin
  	FillChar(CleDoc2, Sizeof(CleDoc2), #0);
  	if TOBPiece.getValue('GP_NATUREPIECEG')='ETU' then
    begin
      CleDoc2.NaturePiece := 'FRC';
      CleDoc2.DatePiece := TOBPiece.GetValue('GP_DATEPIECE');
      CleDoc2.Souche := TOBPiece.GetValue('GP_SOUCHE');
      CleDoc2.NumeroPiece := TOBPiece.GetValue('GP_NUMERO');
      CleDoc2.Indice := TOBPiece.GetValue('GP_INDICEG');
    end else
    begin
    	for Indice := 0 to TOBPiece.detail.count -1 do
      begin
      	if TOBPiece.detail[Indice].getValue('GL_PIECEORIGINE')<>'' then
        begin
        	DecodeRefPiece (TOBpiece.detail[Indice].getValue('GL_PIECEORIGINE'),cledoc2);
          break;
        end;
      end;
      cledoc2.naturepiece := 'FRC';
    end;
    if not ExistePiece(CleDoc2) then
    begin
    	// va bien falloir la créer cette piece
      // on reinit les numero de piece
      cledoc2.Souche := '';
      CleDoc2.NumeroPiece := 0;
      CleDoc2.Indice := 0;
      CleDoc2.DatePiece := V_PGI.DateEntree;

      if WithCreat then
      begin
        EnHt := TOBPiece.GetValue('GP_FACTUREHT') = 'X';
        SaisieContre := TOBPiece.GetValue('GP_SAISIECONTRE') = 'X';
        CreerPieceVide(CleDoc2, TOBPiece.GetValue('GP_TIERS'), TOBPiece.GetValue('GP_AFFAIRE'), TOBPiece.GetValue('GP_ETABLISSEMENT'),
          TOBPiece.GetValue('GP_DOMAINE'), EnHT, SaisieContre);
      	TOBP := TOB.Create ('PIECE',nil,-1);
        PositionnePiece(TOBP,Cledoc2);
        TOBPiece.putValue('GP_PIECEFRAIS',EncodeRefPiece (TOBP));
        TOBPiece.putValue('_NEWFRAIS_','X');
        TOBP.free;
      end else
      begin
      	FillChar(CleDoc2, Sizeof(CleDoc2), #0);
      end;
    end else
    begin
    	// gestion antériorité
      if TOBPiece.getValue ('GP_NATUREPIECEg')<> 'ETU' then  ModeSaisie := TaConsult;
    	exit; // elle est rerouvéeee ouhééé champomy pour tout le monde...
    end;
  end else
  begin
  	DecodeRefPiece (TOBpiece.getValue('GP_PIECEFRAIS'),cledoc2);
    if (not ExistePiece(CleDoc2)) then
    begin
    	if WithCreat then
      begin
      	CleDoc2.NumeroPiece := 0;
      	CleDoc2.DatePiece  := V_PGI.DateEntree;
        EnHt := TOBPiece.GetValue('GP_FACTUREHT') = 'X';
        SaisieContre := TOBPiece.GetValue('GP_SAISIECONTRE') = 'X';
        CreerPieceVide(CleDoc2, TOBPiece.GetValue('GP_TIERS'), TOBPiece.GetValue('GP_AFFAIRE'), TOBPiece.GetValue('GP_ETABLISSEMENT'),
          TOBPiece.GetValue('GP_DOMAINE'), EnHT, SaisieContre);
      	TOBP := TOB.Create ('PIECE',nil,-1);
        PositionnePiece(TOBP,Cledoc2);
        TOBPiece.putValue('GP_PIECEFRAIS',EncodeRefPiece (TOBP));
        TOBPiece.putValue('_NEWFRAIS_','X');
        TOBP.free;
      end else
      begin
    		FillChar(CleDoc2, Sizeof(CleDoc2), #0);
      end;
    end;
  end;
end;

procedure AddPieceFraisAnnul (TOBpieceFrais,TOBPiece : TOB);
var TOBP : TOB;
begin
	TOBP := TOB.create ('UNE PIECE SUP',TOBpieceFrais,-1);
	TOBP.AddChampSupValeur('FRSANULLE',TOBPiece.getValue('GP_PIECEFRAIS'));
  TOBPiece.putValue('GP_PIECEFRAIS','');
end;

procedure AppelFraisDetailBtp (FF : TForm ; TobPiece,TOBOuvrage,TOBPorcs : TOB; InclusStInFg : boolean=false);
var FFact : TFFacture;
  CleDoc2: R_CleDoc;
  retour : TmsFrais;
  Cf,NewCoef: double;
  CreatAutorise : boolean;
  LAction : TActionFiche;
begin
  FFact := TFFacture (FF);
  with FFact do
  begin
    if IsDejaFacture then
    begin
      LAction := TaConsult;
      CreatAutorise := false;
    end else
    begin
  		CreatAutorise := (isPieceGerableFraisDetail (TOBPiece.getValue('GP_NATUREPIECEG'))) and (TFFacture(FF).Action <> taconsult);
    end;
    if (Laction <> Taconsult) then
    begin
      if isPieceGerableFraisDetail(TOBPiece.getValue('GP_NATUREPIECEG')) then LAction := TaModif else LAction := TaConsult;
    end;
		RetrouvePieceFraisBtp (Cledoc2,TOBpiece,LAction,CreatAutorise);
    if (Cledoc2.NumeroPiece <> 0)  then
    begin
    	cf := TOBPiece.getValue('GP_COEFFC');   // ancien coef avant modif
      SauveColList;
      Retour := SaisiePieceFrais(CleDoc2,LAction,TOBPiece.GetValue('GP_TIERS'),TOBPiece.GetValue('GP_AFFAIRE'));
      RestoreColList;
      If retour = TmsFValide then
      begin
        if Laction <> TaConsult then
        begin
          NewCoef := arrondi(GetCoefFC (TOBPiece),9);
          if (NewCoef = 0) and (TOBPiece.getValue('GP_PIECEFRAIS')<>'')then
          begin
            //BTPSupprimePieceFrais (tobpiece);
            AddPieceFraisAnnul (TFFacture(FFact).TOBpieceFraisAnnule,TOBPiece);
          end;
          if (TFFacture(FF).Action=taModif) and (Cf <> NewCoef) then
          begin
            PutValueDetail (TOBPiece,'GP_RECALCULER','X');
      			ReinitPieceForCalc;
          end;

          if FFact.GereDocEnAchat  then
          begin
            AfficheZonePiedAchat;
          end else
          begin
            if isVenteAchat = 'VEN' then AfficheZonePiedStd;
          end;
        end;
      end else if retour = TmsSuppr then
      begin
        AddPieceFraisAnnul (TFFacture(FFact).TOBpieceFraisAnnule,TOBPiece);
        if Cf <> 0 then
        begin
          PutValueDetail (TOBPiece,'GP_RECALCULER','X');
      		ReinitPieceForCalc;
          if FFact.GereDocEnAchat  then
          begin
            AfficheZonePiedAchat;
          end else
          begin
            if isVenteAchat = 'VEN' then AfficheZonePiedStd;
          end;
        end;
      end;
    end;

	end;
end;

procedure GestionDetailOuvrage(FF : Tform ; TobPiece : TOB ; Arow: integer; Acol : integer = -1);
var TOBL: TOB;
	LastRow,LastCol : integer;
  IndiceNomen: integer;
  FFact : TFFacture;
  bc : boolean;
begin
  FFAct := TFFacture (FF);
  TOBL := GetTOBLigne(TOBPiece, ARow); // récupération TOB ligne
  if TOBL = nil then exit;
	With FFact do
  begin
  	if TOBL.getValue('GL_TYPEARTICLE') <> 'OUV' then exit;
    BouttonInVisible;
    GS.BeginUpdate;
    GS.SynEnabled := false;
    if TOBL.getValue('GL_TYPEARTICLE') = 'OUV' then
    begin
      FermerDetailOuvrage (Arow);
      AffichageDetailOuvrage(ARow);
    end;
    GS.EndUpdate;
    GS.Invalidate;
    GS.SynEnabled := true;
    bc := false;
    GSRowEnter(GS, ARow, bc, False);
    BouttonVisible(Arow);
  end;
end;

procedure InformeMetre (FF : Tform ; TOBL : TOB ; Arect : TRect);
var FFact : TFFacture;
  LastBrush, LastPen: Tcolor;
  TheChaine: string;
begin
  FFAct := TFFacture (FF);
  with FFact do
  begin
    if TOBL.NomTable = 'LIGNE' then
       TheChaine := trim(strs(TOBL.GetValue('GL_QTEFACT'), V_PGI.okDecQ));

    LastBrush := GS.Canvas.Brush.Color;
    LastPen := GS.Canvas.Pen.Color;
    GS.Canvas.FillRect(ARect);
    GS.Canvas.Brush.Style := bsSolid;
    GS.Canvas.Brush.Color := clRed;
    GS.Canvas.Pen.Color := clRed;
    GS.Canvas.Polygon([Point(Arect.left, Arect.top), point(Arect.left + 4, Arect.top), Point(Arect.left, Arect.top + 4)]);
    GS.Canvas.Brush.Color := LastBrush;
    GS.Canvas.Pen.Color := LastPen;
    GS.Canvas.TextOut(Arect.Right - GS.canvas.TextWidth(TheChaine) - 3, Arect.Top + 2, TheChaine);
  end;
end;

procedure InformePVBloque (FF : Tform ; TOBL : TOB ; Arect : TRect;TheChaine : string);
var FFact : TFFacture;
  LastBrush, LastPen: Tcolor;
begin
  if TOBL = nil then exit;
  FFAct := TFFacture (FF);
  if (TOBL.GetValue('GL_BLOQUETARIF')<>'X') then exit;
  if TheChaine = '' then exit;
  with FFact do
  begin
    LastBrush := GS.Canvas.Brush.Color;
    LastPen := GS.Canvas.Pen.Color;
    GS.Canvas.FillRect(ARect);
    GS.Canvas.Brush.Style := bsSolid;
    GS.Canvas.Brush.Color := clRed;
    GS.Canvas.Pen.Color := clRed;
    GS.Canvas.Polygon([Point(Arect.left, Arect.top), point(Arect.left + 4, Arect.top), Point(Arect.left, Arect.top + 4)]);
    GS.Canvas.Brush.Color := LastBrush;
    GS.Canvas.Pen.Color := LastPen;
    GS.Canvas.TextOut(Arect.Right - GS.canvas.TextWidth(TheChaine) - 3, Arect.Top + 2, TheChaine);
  end;
end;

procedure InsertLigneInGrilleBtp (FF : TForm;TobPiece,TobTiers,TOBAffaire : TOB; var Arow : integer ; var Acol : integer);
var FFact : TFFacture;
begin
  FFAct := TFFacture (FF);
  with FFact do
  begin
    inc(Arow);
    Acol := GS.FixedCols;
    InsertTOBLigne(TOBPiece, Arow);
    InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, ARow);
    NumeroteLignesGC(GS, TOBPiece);
    if (tModeBlocNotes in SaContexte) then
    begin
      GS.RowCount := TOBPiece.Detail.Count + 2;
      if GS.RowCount < NBRowsInit then GS.RowCount := NBRowsInit;
      GS.height := (GS.rowHeights[1] * (GS.Rowcount-LeRgpBesoin.NbLignes)) + (GS.GridLineWidth * (GS.Rowcount-LeRgpBesoin.NbLignes));
    end else
    begin
      GS.RowCount := TOBPiece.Detail.Count + 1;
      if GS.RowCount < NBRowsInit then GS.RowCount := NBRowsInit;
    end;
  end;
end;

procedure AffecteEcartPaOuv (TOBPiece,TOBL,TOBBases,TOBasesL,TOBOuvrage : TOB ; Delta : double;DEv : Rdevise);
var ArtEcart : string;
    IndiceNomen,IndiceEcart,i : integer;
    TOBNomen,TOBN : TOB;
begin
  ArtEcart := trim(GetParamsoc('SO_BTECARTPMA'));
  if not isExistsArticle(ArtEcart) then
  begin
    PgiBox(TraduireMemoire('L''article d''écart est invalide ou non renseigné#13#10Veuillez le définir'), Traduirememoire('Gestion d''écart'));
    exit;
  end;
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceNomen > 0 then
  begin
    TOBNomen := TOBOuvrage.Detail[IndiceNomen - 1];

    for i := 0 to TOBnomen.detail.count - 1 do
    begin
      TOBN := TOBNomen.detail[I];
      if trim(TOBN.GetValue('BLO_CODEARTICLE')) = ArtEcart then
      begin
        IndiceEcart := I;
        break;
      end;
    end;

    if Arrondi(Delta,V_PGI.OkdecP) <> 0  then
    begin
      if IndiceEcart >= 0 then
      begin
        TOBN := TOBNomen.detail[IndiceEcart];
        TOBN.PutValue('BLO_DPA', TOBN.GetValue('BLO_DPA') + Arrondi(Delta,V_PGI.okdecP));
        TOBN.PutValue('BLO_DPR', TOBN.GetValue('BLO_DPA'));
        if TOBN.GetValue('BLO_DPR')<> 0 then TOBN.putvalue('BLO_COEFMARG', arrondi(TOBN.GetValue('BLO_PUHTDEV')/ TOBN.GetValue('BLO_DPR'),4));
      end else
      begin
        AttribueEcart(TobPiece,TOBL,TobNomen, ArtEcart, Delta, DEV, true, true,True);
      end;
    end;

  end;

end;

procedure MBDetailNomenClickBtp (FF : Tform ; IndiceNomen : integer; TOBaffaire,TobPiece,TobOuvrage,TobArticles,TOBBases,TOBBasesL,TOBL,TOBrepart, TobMetres : TOB ; addicted : boolean; DEV : Rdevise; ModeSaisieAch : boolean;CotraitOk : boolean);
var FFact : TFFacture;
  Valeurs: T_Valeurs;
  TobNomen : TOB;
  //TOBmetres : TOB;
  XP: double;
  TypeARp,EnHt,OkTrait: boolean;
  TypeArticle,RecupPrix : string;
  Taction: TActionFiche;
  MontantPv,Pourcent : double;
//  TOBOUV_O,TOBOUV : TOB;
  LastIndiceLig : Integer;
begin
	RecupPrix := GetInfoParPiece(TOBPIece.GetValue('GP_NATUREPIECEG'), 'GPP_APPELPRIX');
  FFact := TFFacture (FF);
  with FFact do
  begin
    // Ouvrages
//    TOBOUV_O := TOB.Create ('L ANCIENNE DECOMPOSI',nil,-1);
    TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
    TypeArp := (TypeArticle = 'ARP') or (TypeArticle = 'MAR') or (TypeArticle = 'PRE');
    if Action = taconsult then Taction := taconsult else Taction := taModif;
    if (Taction=taModif) and (FFact.ModifAvanc) and (TOBL.GetValue('GL_PIECEPRECEDENTE')<>'') then Taction := TaConsult;
    TOBNomen := TOBOuvrage.Detail[IndiceNomen - 1];
    if TOBNomen = nil then Exit;
    (*
    if Taction <> taConsult then
    begin
      // -- En vue de la déduction de la ligne
      TOBOUV := TOB.Create ('',TOBOUV_O,-1);
      TOBOUV.Dupliquer(TobNomen,True,true); // pour sauvegarder le détail avant modif..
    end;
    *)
    // ------------------------------------
    EnHT := (TOBPiece.GetValue('GP_FACTUREHT') = 'X');
    OkTrait := Entree_LigneOuv(FF,TOBaffaire,TOBNomen, TobArticles, TOBL,TOBRepart,TobMetres, False, 1, 0, DEV, TOBL.GetValue('GL_QUALIFQTEVTE'), valeurs, Taction, Enht, TypeArp, ModeSaisieAch,CotraitOk);
    if (not (action = taconsult)) and (OKtrait) then
    begin
      // Petite gymnastique pour déduire les éléments précédent
      //TOBL.PutValue('GL_INDICENOMEN',0);
      //DeduitLaLignePrec (TOBL,TOBOUV_O);  ZeroLigneMontant (TOBL);
      //TOBL.PutValue('GL_INDICENOMEN',indiceNomen);
      // --------
    	BligneModif := true;
    	if RecupPrix = 'PUH' then
      begin
      	if (TOBL.GetValue('GL_BLOQUETARIF')='X') or ((tModeSaisieBordereau in FFact.SaContexte) and (FFact.OrigineExcel)) then
        begin
        	if (EnHt) and (Valeurs[2] <> TOBL.GetValue('GL_PUHTDEV')) then
          begin
          	MontantPv := TOBL.getValue('GL_PUHTDEV');
            if not (tModeSaisieBordereau in FFact.SaContexte) then
            begin
          		TOBL.PutValue('GL_BLOQUETARIF','-'); // Pour forcer le recalcul des Pv sur le sous détail
            end;
            TraitePrixOuvrage (TOBPiece,TOBL,TobBases,TOBBasesL,TOBOuvrage,EnHt, TOBL.GetValue('GL_PUHTDEV'),Valeurs[2],DEV,false,true);
            if not (tModeSaisieBordereau in FFact.SaContexte) then
            begin
          		TOBL.PutValue('GL_BLOQUETARIF','X');
            end;
          end else if (not EnHT) and (valeurs[3] <> TOBL.GetValue('GL_PUTTCDEV')) then
          BEGIN
          	MontantPv := TOBL.getValue('GL_PUTTCDEV');
            if not (tModeSaisieBordereau in FFact.SaContexte) then
            begin
          		TOBL.PutValue('GL_BLOQUETARIF','-'); // Pour forcer le recalcul des Pv sur le sous détail
            end;
            TraitePrixOuvrage (TOBPiece,TOBL,TobBases,TOBBasesL,TOBOuvrage,EnHt,TOBL.GetValue('GL_PUTTCDEV'),Valeurs[3],DEV,False,True);
            if not (tModeSaisieBordereau in FFact.SaContexte) then
            begin
          		TOBL.PutValue('GL_BLOQUETARIF','X');
            end;
          END;
        end else
        begin
          TOBL.Putvalue('GL_PUHTDEV', valeurs[2]);
          TOBL.Putvalue('GL_PUTTCDEV', valeurs[3]);
        end;
        (*
        if (Valeurs[0] <> TOBL.GetValue('GL_DPA')) then
        begin
           AffecteEcartPaOuv (TOBPiece,TOBL,TOBBases,TOBBasesL,TOBOuvrage,Valeurs[0]-TOBL.GetValue('GL_DPA'),DEv);
        end;
        *)
      end else if RecupPrix = 'DPR' then
      begin
				TOBL.Putvalue('GL_PUHTDEV', valeurs[1]);
        TOBL.Putvalue('GL_PUTTCDEV', valeurs[1]);
      end else if (RecupPrix = 'DPA') or (RecupPrix = 'PAS')  then
      begin
				TOBL.Putvalue('GL_PUHTDEV', valeurs[0]);
        TOBL.Putvalue('GL_PUTTCDEV', valeurs[0]);
      end;
      TOBL.Putvalue('GL_DPA', valeurs[0]);
      TOBL.Putvalue('GL_DPR', valeurs[1]);
      TOBL.Putvalue('GL_PMAP', valeurs[6]);
      TOBL.Putvalue('GL_PMRP', valeurs[7]);
//
      TOBL.PutValue('GL_MONTANTPAFG',valeurs[10]);
      TOBL.PutValue('GL_MONTANTPAFR',valeurs[11]);
      TOBL.PutValue('GL_MONTANTPAFC',valeurs[12]);
      TOBL.PutValue('GL_MONTANTFG',valeurs[13]);
      TOBL.PutValue('GL_MONTANTFR',valeurs[14]);
      TOBL.PutValue('GL_MONTANTFC',valeurs[15]);
//      TOBL.PutValue('GL_MONTANTPA',valeurs[16]);
      TOBL.PutValue('GL_MONTANTPR',valeurs[17]);
//
      TOBL.Putvalue('GL_TPSUNITAIRE', valeurs[9]);
      CalculeMontantsAssocie(TOBL.Getvalue('GL_PUHTDEV'), XP, DEV);
      TOBL.Putvalue('GL_PUHT', XP);
      CalculeMontantsAssocie(TOBL.Getvalue('GL_PUTTCDEV'), XP, DEV);
      TOBL.Putvalue('GL_PUTTC', XP);
      TOBPIece.putvalue('GP_RECALCULER', 'X');
      TOBL.PutValue('GL_RECALCULER', 'X');
      if (not FFact.ModifAvanc) and
      	 (TOBL.fieldExists('BLF_NATUREPIECEG')) and
         (Pos(TOBL.GetValue('BLF_NATUREPIECEG'),'FBT;FBP')>0) then
      begin
        TOBL.Setdouble('BLF_QTESITUATION',TOBL.GetDouble('GL_QTEFACT'));
        TOBL.SetDouble('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_QTESITUATION')*TOBL.GetValue('GL_PUHTDEV'),DEV.decimale));
        TOBL.PutValue('BLF_MTPRODUCTION',TOBL.GetValue('BLF_MTSITUATION'));
        if TOBL.GetValue('BLF_MTMARCHE') <> 0 then
        begin
          Pourcent := Arrondi(TOBL.GetValue('BLF_MTSITUATION')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
        end else
        begin
          Pourcent := Arrondi(100,2);
        end;
        TOBL.PutValue('BLF_POURCENTAVANC',Pourcent);
      end;
      StockeInfoTypeLigne (TOBL,Valeurs);
      positionneCoefMarge (TOBL);
      GS.BeginUpdate;
      TRY
        CalculeLaSaisie(-1, GS.row, True);
        AfficheLaLigne(GS.row);
        MontreInfosLigne(TOBL, nil, nil, nil, INFOSLIGNE, PPInfosLigne);
        if addicted then
        begin
          SuprimeSousDetail(TOBL, TOBOuvrage);
        end else
        begin
          GestionDetailOuvrage(FF,TobPiece,GS.row);
        end;
        if (TOBL.GetString('GL_PIECEPRECEDENTE')='') and (TOBL.FieldExists('BLF_QTESITUATION')) then
        begin
          TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_QTESITUATION')*TOBL.GetValue('GL_PUHTDEV')/TOBL.GEtValue('GL_PRIXPOURQTE'),DEV.decimale));
          TOBL.PutValue('BLF_MTPRODUCTION',TOBL.GetValue('BLF_MTSITUATION'));
        end;
      FINALLY
        GS.EndUpdate;
      END;
      FFact.StCellCur := FFact.GS.Cells [FFact.GS.Col,FFact.GS.row];
    end else
    begin
      if addicted then SuprimeSousDetail(TOBL, TOBOuvrage);
    end;
    // TOBOUV_O.free;
  end;
end;

procedure PositionneColonnesAchatBtp (FF : TForm);
var FFact : TFFacture;
begin
  FFAct := TFFacture (FF);
  with FFact do
  begin
  	if SG_PX > 0 then
    begin
      FFACT.IdentCols[SG_PX].ColName := 'GL_DPA'; // obligatoire pour la gestion des divers dans facture.pas
      SG_PxAch := SG_Px;
    end;
    SG_Px := -1;
    if Sg_Montant > 0 then
    begin
      SG_MontantAch := SG_Montant;
      Sg_Montant := -1;
    end;
  	if SG_PxAch > 0 then GS.Cells[SG_PxAch, 0] := TraduireMemoire('PU Achat');
    if SG_MontantAch > 0 then GS.Cells[SG_MontantAch, 0] := TraduireMemoire('Mt Achat');
    if (Action <> taConsult) and (SG_MontantAch > 0) then GS.ColLengths[SG_MontantAch] := -1;
  end;
end;

procedure PositionneColonnesVenteBtp (FF : TForm);
var FFact : TFFacture;
begin
  // Pas de changement de présentation de la grille
  // si les deux prix unitaires sont déjà présents
  if (SG_PXAch > 0) and (SG_Px > 0) then Exit;

  FFAct := TFFacture (FF);
  with FFact do
  begin
  	if SG_PXACH > 0 then
    begin
    	SG_Px := SG_PxAch;
    end;
    if Sg_PX > 0 then
    begin
    	FFAct.IdentCols[SG_PX].ColName := 'GL_PUHTDEV'; // obligatoire pour la gestion des divers dans facture.pas
    end;
    SG_PxAch := -1;
    if SG_MontantAch > 0 then
    begin
    	SG_Montant := SG_MontantAch;
    end;
    Sg_MontantAch := -1;

  	if SG_Px > 0 then GS.Cells[SG_Px, 0] := TraduireMemoire('PU');
    if SG_Montant > 0 then GS.Cells[SG_Montant, 0] := TraduireMemoire('Montant');
  end;
end;

procedure PositionneColonnesBtp ( FF : TForm ;TOBPiece : TOB);
var FFact : TFFacture;
    TypeFac : string;
begin
  FFAct := TFFacture (FF);
  with FFact do
  begin
  	if Action = TaConsult then exit;
    TypeFac := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
    if ((Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'FBT;FBP')>0) and (TypeFac = 'AVA')) or
      ((Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'DAC;BAC')>0) and (TypeFac = 'DAC')) then
    begin
      // pas d'accès à la référence ni aux prix en modification de situation
      if SG_RefArt <> -1 then GS.ColLengths[SG_RefArt] := -1;
      if SG_Px <> -1 then GS.ColEditables[SG_Px] := false;
      if SG_Montant <> -1 then GS.ColLengths[SG_Montant] := -1;
    end;
    if (OrigineEXCEL) and not(tModeSaisieBordereau in SaContexte) then
    begin
      // pas d'accès à la référence ni aux prix en Etude d'origine EXCEL
      // enleve suite à gestion des sous detail dans la saisie de piece
      //if SG_RefArt <> -1 then GS.ColLengths[SG_RefArt] := -1;
    end;
    if (tModeSaisieBordereau in SaContexte) then
    begin
      // pas d'accès à la quantite (bicose on est sur un bordereau)
      if SG_QF <> -1 then
      begin
      	GS.ColLengths[SG_QF] := -1;
      	GS.ColWidths[SG_QF] := 0;
      end;
    end;
  end;
end;


procedure DefinieColonnesAchatBtp ( FF : TForm);
var FFact : TFFacture;
begin
  FFAct := TFFacture (FF);
  with FFact do
  begin
  	if GS.ListeParam = GetParamsoc('SO_LISTEAVANC') then exit;
    if GP_FACTUREHT.Checked then
    begin
      LesColonnesAch  := FindEtReplace(LesColonnesAch, 'GL_PUHTDEV', 'GL_DPA', False);
      LesColonnesAch := FindEtReplace(LesColonnesAch, 'GL_MONTANTHTDEV', 'GL_MONTANTPA', False);
      LesColonnesAch := FindEtReplace(LesColonnesAch, 'GL_TOTALHTDEV', 'GL_MONTANTPA', False);
    end
    else
    begin
      LesColonnesAch := FindEtReplace(LesColonnesAch, 'GL_PUTTCDEV', 'GL_DPA', False);
      LesColonnesAch := FindEtReplace(LesColonnesAch, 'GL_MONTANTTTCDEV', 'GL_MONTANTPA', False);
      LesColonnesAch := FindEtReplace(LesColonnesAch, 'GL_TOTALTTCDEV', 'GL_MONTANTPA', False);
    end;
  end;
end;

procedure PositionneDocAchatBtp ( FF : TForm);
var FFact : TFFacture;
begin
  FFAct := TFFacture (FF);
  with FFact do
  begin
    LesColonnes := LesColonnesAch;
    GS.titres[0] := LesColonnes;
  end;
end;


procedure PositionneDocVenteBtp ( FF : TForm);
var FFact : TFFacture;
begin

  FFAct := TFFacture (FF);
  with FFact do
  begin
  	LesColonnes := LesColonnesVte;
    if not GP_FACTUREHT.Checked then
    begin
      LesColonnes := FindEtReplace(LesColonnes, 'GL_PUHTDEV', 'GL_PUTTCDEV', False);
      LesColonnes := FindEtReplace(LesColonnes, 'GL_TOTALHTDEV', 'GL_TOTALTTCDEV', False);
      LesColonnes := FindEtReplace(LesColonnes, 'GL_MONTANTHTDEV', 'GL_MONTANTTTCDEV', False);
      // ajout 02/08/2001
      LesColonnes := FindEtReplace(LesColonnes, 'GL_PUHTNETDEV', 'GL_PUTTCNETDEV', False);
      GS.Titres[0] := LesColonnes;
      (*
      LesColonnesVte := lesColonnes;
      LesColonnesAch := lesColonnes;
      DefinieColonnesAchatBtp(self);
      *)
    end;
    if FFact.LaPieceCourante.GetString('GP_VENTEACHAT')='ACH' then
      LesColonnes := FindEtReplace(LesColonnes, 'GL_QUALIFQTEVTE', 'GL_QUALIFQTEACH', False)
    else
      LesColonnes := FindEtReplace(LesColonnes, 'GL_QUALIFQTEACH', 'GL_QUALIFQTEVTE', False);
    GS.titres[0] := LesColonnes;
  end;
end;

procedure TentativeAffectationQtive (TOBL,TOBOuvrage : TOB;EnPa,EnPr,EnHt : boolean;PuSaisi : double ; var SumLoc: double ; DEV:rdevise; ForcePaEqualPv : Boolean=false);
var Indice : integer;
		indPa,IndPr,IndPVHT,IndPVHTDEv : integer;
    restitue,restituePr,EnCoursPr,EnCours ,MontantLoc,MontantLocPr, LocSumLoc,LocSumLocPr,LastMontantLoc,LastMontantLocPr: double;
    TOBO,TOBPiece : TOB;
    CoefpaPr,CoefPrPV,SumPr : double;
    CoefFgFixe : boolean;
    QTe,QteDudetail : double;
begin
	TOBPiece := TOBL.Parent;
  CoefFgFixe := false;
  if (TOBPiece <> nil) and
  	 (TOBPiece.FieldExists('COEFFGFORCE')) and
     (TOBPiece.getValue('COEFFGFORCE')<> 0) then CoefFgFixe := true;
  SumPr := 0;
	Restitue := Arrondi(PuSaisi-SumLoc,V_PGI.Okdecp);
  restituePr := 0;
	SumLoc := 0;
  IndPa := TOBOuvrage.detail[0].GetNumChamp ('BLO_DPA');
  IndPr := TOBOuvrage.detail[0].GetNumChamp ('BLO_DPR');
  if EnHt then
  begin
    IndPVHT := TOBOuvrage.detail[0].GetNumChamp ('BLO_PUHT');
    IndPVHTDEV := TOBOuvrage.detail[0].GetNumChamp ('BLO_PUHTDEV');
  end else
  begin
    IndPVHT := TOBOuvrage.detail[0].GetNumChamp ('BLO_PUTTC');
    IndPVHTDEV := TOBOuvrage.detail[0].GetNumChamp ('BLO_PUTTCDEV');
  end;

  for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
  	TOBO := TOBOuvrage.detail[Indice];
    if Indice = 0 then QteDuDetail := TOBO.GetValue('BLO_QTEDUDETAIL');
    if QteDuDetail = 0 then QteDudetail := 1;
    if TOBO.GetValue('BLO_QTEFACT') = 0 then continue;
    Qte := TOBO.GetValue('BLO_QTEFACT');
    EnCours :=  Arrondi (Restitue/ Qte,V_PGI.OKdecP);
    EnCoursPr :=  Arrondi (RestituePr/Qte,V_PGI.OKdecP);
    if (TOBO.GetValue('BLO_TYPEARTICLE')='OUV') and
       (not Isvariante(TOBO)) and
       (EnCOurs<>0) then
    begin
      if EnPa then LastMontantLoc := TOBO.getValue('BLO_MONTANTPA')/QteDuDetail
      else if EnHt then LastMontantLoc := arrondi(TOBO.getValue('BLO_MONTANTHTDEV')/qteDudetail,V_pGI.okdecP)
      else LastMontantLoc := Arrondi(TOBO.getValue('BLO_MONTANTTTCDEV')/QteDuDetail,V_pGI.okdecP);
      if TOBO.getValeur(indPa) > 0 then CoefpaPr := TOBO.getValeur(indPr) / TOBO.getValeur(indPa)
                                   else CoefpaPr := 1;
      if TOBO.getValeur(indPr) > 0 then CoefprPv := TOBO.getValeur(indPvHtDev) / TOBO.getValeur(indpr)
      														 else CoefprPv := 1;

      LastMontantLocPr := Arrondi(TOBO.getValue('BLO_MONTANTPR')/QteDudetail,V_PGI.okdecP);
      LocSumLoc := 0;
      LocSumLocPr := 0;

    	EssayeAffecterEcartQte (TOBL,TOBO,Encours,EncoursPr,LocSumLoc,LocSumLocPr,IndPa,IndPr,IndPVHtdev,IndPvHt,EnPa,EnPR,EnHt,DEV,True);

      if ENPA then
      begin
        TOBO.Putvaleur(IndPa, LocSumLoc);
        TOBO.PutValue('GA_PAHT', TOBO.GetValue('BLO_DPA'));
        if not CoefFgFixe then TOBO.Putvaleur(IndPr, Arrondi(LocSumLoc * CoefPaPr, V_PGI.okDecP));
        if TOBL.GetValue('GL_BLOQUETARIF') = '-' then
        begin
          TOBO.Putvaleur(IndPvHtDev, Arrondi(TOBO.getValeur(IndPr) * CoefPrPv, V_PGI.okdecP));
          TOBO.Putvaleur(IndPvHt, devisetopivotEx(TOBO.Getvaleur(IndPvHtDev), DEV.Taux, DEV.quotite, V_Pgi.okdecP));
          TOBO.PutValue('GA_PVHT', TOBO.Getvaleur(IndPvHtdev));
          if TOBO.GetValue('BLO_TYPEARTICLE') = 'PRE' then
            TOBO.PutValue('GF_PRIXUNITAIRE', TOBO.Getvaleur(IndPvHtdev));
        end;
      end else
      begin
        if not CoefFgFixe then TOBO.Putvaleur(Indpr, LocSumLocPr);
        if TOBL.GetValue('GL_BLOQUETARIF') = '-' then
        begin
          TOBO.Putvaleur(IndPvHtDev, LocSumLoc);
          TOBO.Putvaleur(IndPvHt, devisetopivotEx(LocSumLoc, DEV.Taux, DEV.quotite, V_PGI.okdecP));
          if TOBO.Getvalue('BLO_DPR') <> 0 then TOBO.Putvalue('BLO_COEFFG',Arrondi(TOBO.Getvaleur(IndPvHt)/TOBO.Getvalue('BLO_DPR'),4));
          if ForcePaEqualPv then
          begin
            TOBO.SetDouble('BLO_DPA',LocSumLoc);
            TOBO.SetDouble('BLO_DPR',LocSumLoc);
            TOBO.SetDouble('BLO_COEFFG',0);
            TOBO.SetDouble('BLO_COEFFC',0);
            TOBO.SetDouble('BLO_COEFFR',0);
            TOBO.SetDouble('BLO_COEFMARG',1);
          end;
        end;
      end;
      CalculMontantHtDevLigOuv(TOBO, DEV);
      if EnPa then MontantLoc := TOBO.getValue('BLO_MONTANTPA')
      else if EnHt then MontantLoc := TOBO.getValue('BLO_MONTANTHTDEV')
      else MontantLoc := TOBO.getValue('BLO_MONTANTTTCDEV');

      MontantLocPr := TOBO.getValue('BLO_MONTANTPR');

      restitue := ARRONDI(restitue - (MontantLoc - lastMontantLoc),V_PGI.OkDecP);
      restituePr := ARRONDI(restituePr - (MontantLocPr - LastMontantLocPr),V_PGI.OkDecP);
    end else
    begin
      if (TOBO.GEtValue('BLO_TYPEARTICLE')<> 'OUV' ) and
         (TOBO.GEtValue('BLO_TYPEARTICLE')<> 'POU') and (not IsVariante(TOBO)) and
         (EnCours <> 0) then
      begin
        if EnPa then LastMontantLoc := TOBO.getValue('BLO_MONTANTPA')/QteDuDetail
        else if EnHt then LastMontantLoc := arrondi(TOBO.getValue('BLO_MONTANTHTDEV')/QteDuDetail,V_PGI.okdecP)
        else LastMontantLoc := Arrondi(TOBO.getValue('BLO_MONTANTTTCDEV')/QteDuDetail,V_PGI.okdecP);

        LastMontantLocPr := Arrondi(TOBO.getValue('BLO_MONTANTPR') / QteDuDetail,V_PGI.okdecP);
        AffecteDifferenceSurLigne (TOBL,TOBO,EnCours,EnCoursPr,IndPa,IndPr,IndPVHtDev,IndPVHt,EnPa,EnPr,EnHt,DEV,ForcePaEqualPv);
        CalculMontantHtDevLigOuv(TOBO, DEV);
        if EnPa then MontantLoc := Arrondi(TOBO.getValue('BLO_MONTANTPA')/QteDuDetail,V_PGI.okdecP)
        else if EnHt then MontantLoc := Arrondi(TOBO.getValue('BLO_MONTANTHTDEV')/QteDuDetail,V_PGI.okdecP)
        else MontantLoc := Arrondi(TOBO.getValue('BLO_MONTANTTTCDEV')/QteDuDetail,V_PGI.okdecP);

        MontantLocPr := Arrondi(TOBO.getValue('BLO_MONTANTPR')/QteDuDetail,V_PGI.okdecP);

        restitue := ARRONDI(restitue - (MontantLoc - lastMontantLoc),V_PGI.OkDecP);
        restituePr := ARRONDI(restituePr - (MontantLocPr - LastMontantLocPr),V_PGI.okdecP);

      end else
      begin
        if EnPa then MontantLoc := Arrondi(TOBO.getValue('BLO_MONTANTPA')/QteDuDetail,V_PGI.okdecP)
        else if EnHt then MontantLoc := Arrondi(TOBO.getValue('BLO_MONTANTHTDEV')/QteDuDetail,V_PGI.okdecP)
        else MontantLoc := Arrondi(TOBO.getValue('BLO_MONTANTTTCDEV')/QteDuDetail,V_PGI.okdecP);

        MontantLocPr := Arrondi(TOBO.getValue('BLO_MONTANTPR')/QteDuDetail,V_PGI.okdecP);
      end;
    end;
    // --
    SumLoc := SumLoc + MontantLoc;
    SumPr := SumPr + MontantLocPr;
  end;
  SumLoc := ARRONDI(SumLoc,V_PGI.OkDecP);
  SumPr := ARRONDI(SumPr,V_PGI.OkDecP);
end;

procedure  RecupValopose (TOBOUV : TOB);
begin
	GetValoDetail (TOBOUV.detail[0]);
end;

function TraitePrixOuvrage(TOBPiece, TOBL, TOBBases, TOBBasesL,TOBOuvrage: TOB; EnHt: boolean; PuFix,PuReference: double; DEV: Rdevise; EnPA: boolean = false; FromReference : boolean=false; AffecteEcartLig : Boolean=true; ForcePaEqualPv : Boolean=false): boolean;
var
  IndiceNomen, i, IndPVHTDEv: Integer;
  TobNomen, TOBN: TOB;
  PURef, PUSaisi: double;
  EnPr : boolean;
  PrCalc,CoefPrPv,MtSit : double;
  resultOk: boolean;
  ArtEcart: string;
  IndiceEcart: integer;
begin
  IndiceEcart := -1;
  ArtEcart := trim(GetParamsoc('SO_BTECARTPMA'));
  if not isExistsArticle(ArtEcart) then
  begin
    PgiBox(TraduireMemoire('L''article d''écart est invalide ou non renseigné#13#10Veuillez le définir'), Traduirememoire('Gestion d''écart'));
    exit;
  end;
  ResultOk := true;
  if Fromreference then
  begin
  	PuRef := PUReference;
  end else
  begin
    if EnPa then PuRef := TOBL.GetValue('GL_DPA') else
    if EnHt then PuRef := TOBL.GetValue('GL_PUHTDEV')
    else PuRef := TOBL.GetValue('GL_PUTTCDEV');
  end;
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceNomen > 0 then
  begin
    TOBNomen := TOBOuvrage.Detail[IndiceNomen - 1];

    for i := 0 to TOBnomen.detail.count - 1 do
    begin
      TOBN := TOBNomen.detail[I];
      if trim(TOBN.GetValue('BLO_CODEARTICLE')) = ArtEcart then
      begin
        IndiceEcart := I;
        break;
      end;
    end;

    // boucle pour essayer d'affecter complètement la valeur saisie
    // on fait 5 tentatives maxi
    PuSaisi := PuFix;
    i := 1;
    repeat
      if not ResultOk then break;
      resultOk := ReajusteMontantOuvrage(nil,TOBPiece,TOBL, TOBNomen, PuRef, PrCalc, PuFix, DEV, EnHt, false, false, EnPA,ForcePaEqualPv);
      if ResultOk then
      begin
        if ARRONDI(PuFix,V_PGI.OkDecP) = ARRONDI(PuSaisi,V_PGI.OkDecP) then i := 4
        else
        begin
          Inc(i);
          if i < 4 then
          begin
            PuRef := PuFix;
            PuFix := PuSaisi;
          end;
        end;
      end;
    until (i = 4);
    if ARRONDI(PuFix,V_PGI.OkDecP) <> Arrondi(PuSaisi,V_PGI.OkDecP) then
    begin
    	EnPr := false;
      TentativeAffectationQtive (TOBL,TOBNomen,EnPa,EnPr,EnHt,PuSaisi,PuFix,DEV,ForcePaEqualPv);
    end;
    if AffecteEcartLig then
    begin
      if Arrondi(PuFix - PuSaisi,V_PGI.OkdecP) <> 0  then
      begin
        if IndiceEcart >= 0 then
        begin
          TOBN := TOBNomen.detail[IndiceEcart];
          if EnPa then
          begin
            TOBN.PutValue('BLO_DPA', TOBN.GetValue('BLO_DPA') + PuSaisi - PuFix);
            TOBN.PutValue('BLO_DPR', TOBN.GetValue('BLO_DPR') + PuSaisi - PuFix);
          end;
          if (TOBL.GetValue('GL_BLOQUETARIF')='-')  then
          begin
            if (EnHt) then
            begin
              TOBN.PutValue('BLO_PUHTDEV', Arrondi(TOBN.GetValue('BLO_PUHTDEV') + PuSaisi - PuFix,V_PGI.okdecP));
            end else
            begin
              TOBN.PutValue('BLO_PUTTCDEV', Arrondi(TOBN.GetValue('BLO_PUTTCDEV') + PuSaisi - PuFix,V_PGI.okdecP));
            end;
          end;
          TOBN.putvalue('BLO_PUHT', DeviseToPivotEx(TOBN.GetValue('BLO_PUHTDEV'), DEV.Taux, DEV.quotite, V_PGI.okdecP));
          if (not EnPA) and (ForcePaEqualPv) then
          begin
            TOBN.SetDouble('BLO_DPA',TOBN.GetDouble('BLO_PUHT') );
            TOBN.SetDouble('BLO_DPR',TOBN.GetDouble('BLO_DPA') );
            TOBN.SetDouble('BLO_COEFFG',0 );
            TOBN.SetDouble('BLO_COEFFC',0 );
            TOBN.SetDouble('BLO_COEFFR',0 );
            TOBN.SetDouble('BLO_COEFMARG',1);
          end;
          TOBN.putvalue('BLO_PUHTBASE', TOBN.getValue('BLO_PUHT'));
          TOBN.putvalue('ANCPA', TOBN.Getvalue('BLO_DPA'));
          TOBN.putvalue('ANCPR', TOBN.Getvalue('BLO_DPR'));
          if TOBN.GetValue('BLO_DPR')<> 0 then TOBN.putvalue('BLO_COEFMARG', arrondi(TOBN.GetValue('BLO_PUHTDEV')/ TOBN.GetValue('BLO_DPR'),4));
        end else
        begin
          AttribueEcart(TobPiece,TOBL,TobNomen, ArtEcart, ARRONDI(PuSaisi - PuFix,V_PGI.OkDecP), DEV, ENHt, EnPA,False,0,'',ForcePaEqualPv);
        end;
        PuFix := PuSaisi;
      end;
    end;
    if enPa then
    begin
    	if EnHt then IndPvHtdev := TOBL.GetNumChamp ('GL_PUHTDEV')
      				else IndPVHtDev := TOBL.GetNumChamp ('GL_PUTTCDEV');
      if TOBL.getValue('GL_DPR') > 0 then CoefprPv := TOBL.getValeur(indPvHtDev) / TOBL.getValue('GL_DPR')
                                  	 else CoefprPv := 1;
      TOBL.PutValue('GL_DPA', PuFix);
      //      TOBL.PutValue('GL_DPR', Arrondi(PuFix * CoefPaPr,V_PGI.OkdecP));
      TOBL.PutValue('GL_DPR', Arrondi(PrCalc, V_PGI.OkdecP));
			if (TOBL.GetValue('GL_BLOQUETARIF')='-') then
      begin
      	TOBL.PutValue('GL_PUHTDEV', Arrondi(TOBL.GetValue('GL_DPR') * CoefPrPv, V_PGI.OkdecP));
      end;
      TOBL.PutValue('GL_RECALCULER', 'X');
    end else
    if EnHt then
    begin
			if (TOBL.GetValue('GL_BLOQUETARIF')='-') then
      begin
        TOBL.PutValue('GL_PUHTDEV', PuFix);
        if (ForcePaEqualPv) then
        begin
          TOBL.SetDouble('GL_DPA',TOBL.GetDouble('GL_PUHTDEV') );
          TOBL.SetDouble('GL_DPR',TOBL.GetDouble('BLO_DPA') );
          TOBL.SetDouble('GL_COEFFG',0 );
          TOBL.SetDouble('GL_COEFFC',0 );
          TOBL.SetDouble('GL_COEFFR',0 );
          TOBL.SetDouble('GL_COEFMARG',1);
        end;
      end;
      TOBL.PutValue('GL_RECALCULER', 'X');
    end else
    begin
			if (TOBL.GetValue('GL_BLOQUETARIF')='-') then
      		TOBL.PutValue('GL_PUTTCDEV', PuFix);
      TOBL.PutValue('GL_RECALCULER', 'X');
    end;
    if TOBL.getValue('GL_TYPEARTICLE')='ARP' then
    begin
      RecupValopose (TOBNomen);
    end;

  end else
  begin
  	TOBL.PutValue('GL_RECALCULER', 'X');
    if EnPa then TOBL.PutValue('GL_DPA',PuFix) else
    if EnHt then
    begin
			if (TOBL.GetValue('GL_BLOQUETARIF')='-') then
      begin
        TOBL.PutValue('GL_PUHTDEV',PuFix);
        if (ForcePaEqualPv) then
        begin
          TOBL.SetDouble('GL_DPA',TOBL.GetDouble('GL_PUHTDEV') );
          TOBL.SetDouble('GL_DPR',TOBL.GetDouble('BLO_DPA') );
          TOBL.SetDouble('GL_COEFFG',0 );
          TOBL.SetDouble('GL_COEFFC',0 );
          TOBL.SetDouble('GL_COEFFR',0 );
          TOBL.SetDouble('GL_COEFMARG',1);
        end;
      end;
    end else
    begin
			if (TOBL.GetValue('GL_BLOQUETARIF')='-') then
    			TOBL.PutValue('GL_PUTTCDEV',PuFix);
    end;
    //
    //
  end;
  // cas de la modification du montant facture sur une ligne ajouté manuellement dans la facture
  if (TOBL.GetString('GL_PIECEPRECEDENTE')='') and
     (TOBL.FieldExists('BLF_QTESITUATION')) then
  begin
    for i := 0 to TOBnomen.detail.count - 1 do
    begin
      TOBN := TOBNomen.detail[I];
      MtSit := arrondi(TOBN.getValue('BLO_QTEFACT')*TOBL.getvalue('GL_QTEFACT') /TOBN.getValue('BLO_QTEDUDETAIL')*TOBN.getValue('BLO_PUHTDEV'),DEV.Decimale);
      TOBN.PutValue('BLF_MTSITUATION',MtSit);
      TOBN.PutValue('BLF_MTPRODUCTION',TOBN.GetValue('BLF_MTSITUATION'));
    end;
  end;
  result := (TOBL.GetValue('GL_RECALCULER')= 'X');
end;

function ZoomOuChoixArtLib(FF : TForm ; TOBPiece,TOBArticles,TOBTiers,TOBAffaire : TOB ;ACol, ARow: integer): Boolean;
var OkArt: boolean;
  TOBArt, TOBL: TOB;
  ReferenceSav, UniteSav, RefUnique, LibSav,RefTiers,RefSpigao,RefArt: string;
  QteSav,PUSav: Double;
  FFact : TFFacture;
  PUFixe : boolean;
  QQ : TQuery;
  RowRef : integer;
  LigXls : integer;
begin
  FFact := TFFActure( FF);
  result := false;
  LigXls := 0;
  With FFact do
  begin
    if Action <> taConsult then
    begin
			BouttonInVisible;
      //
    	RowRef := Arow;
      TOBL := GetTOBLigne(TOBPiece, ARow);
      if TOBL = nil then exit;

      if TOBL.GetValue('GL_QTEFACT') = 0 then exit;
      LibSav := TOBL.GetValue('GL_LIBELLE');
      if (SG_REFTiers <> -1) and (tModeSaisieBordereau in FFact.SaContexte ) then
      begin
      	RefTiers := TOBL.GetValue('GL_REFARTTIERS');
      end;
      PUFixe := (TOBL.GEtValue('GL_BLOQUETARIF')='X');
      if PuFixe then PuSav := TOBL.GetValue('GL_PUHTDEV')
      					else PuSav := 0;
      RefSpigao := TOBL.GetValue('GL_IDSPIGAO');  // sauvegarde de l'identifiant de ligne SPIGAO
      if IsFromExcel(TOBL) then LigXls := TOBL.GetValue('BLX_NUMLIGXLS');
      ReferenceSav := GS.Cells[SG_RefArt, ARow];
      QteSav := Valeur(GS.Cells[SG_QF, ARow]);
      UniteSav := GS.Cells[SG_Unit, ARow];
      OkArt := Lexical_RechArt(GS, HTitres.Mess[1], CleDoc.NaturePiece, GP_DOMAINE.Value, '',RefArt);
      GS.Col :=SG_Lib;
      if not OkArt then
      begin
        if GS.Row > TOBPiece.Detail.count then result := False; // Cas où on stoppe la recherche globale
        GS.Row := ARow;
        GS.Cells[SG_Lib, ARow] := LibSav;
      	TOBL := GetTOBLigne(TOBPiece, ARow);
        TOBL.PutValue('GL_LIBELLE', LibSav);
      end else
      begin
        result := true;
        //
        TOBArt := CreerTOBArt(TOBArticles);
        QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                       'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                       'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+RefArt+'"',true,-1, '', True);
        if not QQ.eof then TOBArt.SelectDB('',QQ);
        ferme (QQ);
        InitChampsSupArticle (TOBArt);

        if TOBArt.Getvalue('GA_STATUTART') = 'GEN' then
        begin
          RefUnique := TOBArt.Getvalue('GA_ARTICLE');
          RefUnique := SelectUneDimension(RefUnique);
          TOBArt.InitValeurs(false);
          QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                         'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                         'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+RefUnique+'"',true,-1, '', True);
          if not QQ.eof then TOBArt.SelectDB('',QQ);
          ferme (QQ);
        end else
        begin
          RefUnique := RefArt;
        end;

        CalculePrixArticle(TOBArt);
        if TOBPiece.GetValue('GP_NATUREPIECEG') = VH_GC.AFNatProposition then
          GS.Cells[SG_RefArt, ARow] := ReferenceSav;
        //
        if not ArticleAutorise(TOBPiece, TOBArticles, CleDoc.NaturePiece, ARow) then
        begin
          HPiece.Execute(2, Caption, '');
          VideCodesLigne(TOBPiece, ARow);
          InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, ARow);
          GS.Cells[SG_Lib, ARow] := LibSav;
          TOBL.PutValue('GL_LIBELLE', LibSav);
        end else
        begin
      		deduitLaLigne(TOBL);ZeroLigneMontant (TOBL);
          if IsOuvrage (TOBL) then
          begin
            SuppressionDetailOuvrage (TOBL);
          end;
        	CodesArtToCodesLigne(TOBArt, ARow);
          if QTeSav = 0 then QteSav := 1;

          if (Pufixe) then
          begin
            if PuSav <>0  then TOBL.PutValue('GL_PUHTDEV',PuSav);
            TOBL.PutValue('GL_BLOQUETARIF','X');
            TOBL.PutValue('GL_RECALCULER', 'X');
          end;

          UpdateArtLigne(ARow, False, False, True,QteSav);

          if (TOBPiece.GetValue('GP_NATUREPIECEG') = VH_GC.AFNatProposition) or
           	 (TOBPiece.GetValue('GP_NATUREPIECEG') = GetParamSoc('SO_BTNATBORDEREAUX')) or (Pufixe) then
          begin
            TOBL.PutValue('GL_LIBELLE', LibSav);
            GS.Cells[SG_Lib, ARow] := LibSav;
            TOBL.PutValue('GL_QTEFACT', QteSav);
            TOBL.PutValue('GL_QUALIFQTEVTE', UniteSav);
            if (SG_REFTiers <> -1) and (tModeSaisieBordereau in FFact.SaContexte ) then
            begin
			      	TOBL.PutValue('GL_REFARTTIERS',RefTiers);
            end;
            if (Pufixe) then
            begin
			      	TOBL.PutValue('GL_BLOQUETARIF','X');
            end;
            if PuSav <>0  then TOBL.PutValue('GL_PUHTDEV',PuSav);
            AfficheLaLigne(Arow);
          end;
          //
          TOBL.PutValue('GL_RECALCULER', 'X');
          TOBPiece.PutValue('GP_RECALCULER', 'X');
          //
          if IsOuvrage (TOBL) then
          begin
            AffichageDetailOuvrage (RowRef);
          end;
        end;
        TOBL.PutValue('GL_IDSPIGAO',RefSpigao);  // restauration de l'identifiant de ligne SPIGAO
        if LigXls <> 0 then TOBL.PutValue('BLX_NUMLIGXLS',LigXls);
        AfficheLaLigne(Arow);
        GS.InvalidateRow(Arow);
        GoToLigne(Arow,Acol);
        GereEnabled(Arow);
      end;
    end;
  end;
end;

function IsOuvrage (TOBL : TOB) : boolean;
var prefixe : string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  if (TOBL.GetValue(prefixe+'_TYPEARTICLE')='OUV') or (TOBL.GetValue(prefixe+'_TYPEARTICLE')='ARP') then result := true;
end;

function ControleReceptionFournisseurBTP (TOBV : TOB; cledoc : r_cledoc;QteSais : double;var PUA : double;TheAction : TactionFiche) : boolean;
var mtAchat : double;
		QteReceptionne,FUA,FUV,FUS,CoefUSUV : double;
    MtReceptionne : Double;
    OK_ReliquatMt : Boolean;
    QQ : TQuery;
    Requete : string;
    TOBLIGNes,TOBL : TOB;
    indice : integer;
begin
	result := true;
  CoefUSUV := TOBV.GetDouble('GL_COEFCONVQTEVTE');
	TOBLignes := TOB.Create ('LES LIGNS',nil,-1);
  PUA := 0;
  QteReceptionne := 0;
  MtReceptionne  := 0;
  MtAchat := 0;
  TRY
    if TOBV <> nil then FUS := RatioMesure('PIE', TOBV.GetValue('GL_QUALIFQTESTO'));
    if TOBV <> nil then FUV := RatioMesure('PIE', TOBV.GetValue('GL_QUALIFQTEVTE'));
//		requete := 'SELECT GL_QTERESTE,GL_QUALIFQTEACH,GL_PUHT,GL_COEFCONVQTE FROM LIGNE WHERE' + WherePiece (cledoc,ttdligne,true,true)+ ' AND GL_QTERESTE >0' ;
		requete := 'SELECT GL_QTEFACT,GL_QTERESTE,GL_MTRESTE,GL_QUALIFQTEACH,GL_PUHT,GL_COEFCONVQTE,GL_MTRESTE FROM LIGNE WHERE' + WherePiece (cledoc,ttdligne,true,true); // + ' AND GL_QTERESTE >0' ;
    QQ := OpenSql (requete,true,-1, '', True);
    if not QQ.eof then
    begin
    	TOBLignes.LoadDetailDB ('LIGNE','','',QQ,false);
      for Indice := 0 to TOBlignes.detail.count -1 do
      begin
        TOBL := TOBLignes.detail[Indice];
        OK_ReliquatMt := CtrlOkReliquat(TOBL, 'GL');
    		if TOBL <> nil then FUA := RatioMesure('PIE', TOBL.GetValue('GL_QUALIFQTEACH'));
        if TOBL.GEtValue('GL_QTERESTE') > 0 then
        begin
          // cas standard ou la reception n'est pas facturée
          if TOBL.getValue('GL_COEFCONVQTE')<> 0 then
          begin
            QteReceptionne := QteReceptionne + (TOBL.GEtValue('GL_QTERESTE') * TOBL.GEtValue('GL_COEFCONVQTE'));
          end else
          begin
            QteReceptionne := QteReceptionne + TOBL.GEtValue('GL_QTERESTE') / FUA  * FUS;
          end;
          mtAchat := MtAchat + (TOBL.GEtValue('GL_QTERESTE') * TOBL.GetValue('GL_PUHT'));
        end
        else
        begin
          // cas ou la reception est facturée
					if TOBL.getValue('GL_COEFCONVQTE')<> 0 then
          begin
            QteReceptionne := QteReceptionne + (TOBL.GEtValue('GL_QTEFACT') * TOBL.GEtValue('GL_COEFCONVQTE'));
          end else
          begin
            QteReceptionne := QteReceptionne + TOBL.GEtValue('GL_QTEFACT') / FUA  * FUS;
          end;
          mtAchat := MtAchat + (TOBL.GEtValue('GL_QTEFACT') * TOBL.GetValue('GL_PUHT'));
        end;
        if OK_ReliquatMt then MtReceptionne := MtReceptionne + TOBL.GEtValue('GL_MTRESTE');
      end;
    end;
    ferme (QQ);
    if (CoefUSUV <> 0) then
    begin
			Qtereceptionne := QteReceptionne * CoefUSUV;
    end else if (FUS*FUV) <> 0 then
    begin
    	Qtereceptionne := QteReceptionne / (FUV*FUS);
    end;
    if QteSais > QteReceptionne then
    begin
      PGIBox (TraduireMemoire('Quantité non réceptionnée'),TraduireMemoire('Controle quantitatif'));
      result := false;
    end else
    begin
      if QteReceptionne > 0 then PUA := arrondi(MTachat/QteReceptionne,V_PGI.OkDecP)
                            else PUA := 0;
    end;
  finally
  	TOBLIgnes.free;
  END;
end;

function ControleQteBTP (TOBL : TOB; RefOrigine : string; QteSais,QteResteInit : double;var PUA : double; TheAction : TactionFiche) : boolean;
var {QteSais,}QteRecep,QTeLivre : double;
    MtRecep : Double;
    MTLivre : Double;
    QQ : TQuery;
    requete : string;
    MtAchat : double;
    cledoc : r_cledoc;
begin
  result    := true;
  QteRecep  := 0;
  QteLivre  := 0;
  MtLivre   := 0;
  MtAchat   := 0;
  PUA       := 0;
//  if RefOrigine = '' then exit;
	if RefOrigine = '' then
  begin
		if TOBL.GEtValue('GL_PIECEORIGINE') <> '' then
    begin
       DecodeRefPiece (TOBL.GEtValue('GL_PIECEORIGINE'),cledoc);
       if Pos(Cledoc.NaturePiece , GetPieceAchat (false,false,false,true)) > 0 then
       begin
         // cas d'une livraison directe par le fournisseur sans besoin de chantier
         result := ControleReceptionFournisseurBTP (TOBL,cledoc,QteSais,PUA,TheAction);
         exit;
       end;
    end else
    begin
    	result := false;
      if PgiAsk(TraduireMemoire('Cette marchandise n''a pas été réceptionnée.#13#10Confirmez-vous la livraison ?'))=MrYes then
      begin
  			TOBL.PutValue('_RUPTURE_AUTH_','X');
      	result := true;
      end ;
      // PGIBox (TraduireMemoire('Quantité non réceptionnée'),TraduireMemoire('Controle quantitatif'));
      exit;
    end;
  end;
  if Pos(TOBL.GetValue('GL_NATUREPIECEG'),GetNaturePieceBLC(false)) <= 0 then exit;
  if IsPrestationInterne (TOBL.GetValue('GL_ARTICLE')) then exit;
//  QteSais := TOBL.GetValue('GL_QTESTOCK');
//  requete := 'SELECT SUM(GL_QTERESTE) AS QTELIVRE FROM LIGNE WHERE GL_PIECEORIGINE="'+RefOrigine+
  requete := 'SELECT SUM(GL_QTERESTE) AS QTELIVRE, SUM(GL_MTRESTE) AS MTLIVRE FROM LIGNE WHERE GL_PIECEPRECEDENTE="'+RefOrigine+
             '" AND GL_NATUREPIECEG IN ("'+GetNaturePieceBLC+'") AND (GL_QTERESTE <> 0 OR GL_MTRESTE <> 0)';
  QQ := OpenSql (requete,true,-1, '', True);
  TRY
    if not QQ.eof then
    begin
      QteLivre  := QQ.findField('QTELIVRE').AsFloat;
      MtLivre   := QQ.findField('MTLIVRE').AsFloat;
    end;
  FINALLY
    Ferme (QQ);
  END;

  if TheAction = TaModif then QteLivre := QteLivre - QteResteInit;

  requete := 'SELECT SUM(GL_QTERESTE) AS QTERECEP,SUM (GL_QTESTOCK*GL_PUHT) AS MTACHAT, SUM(GL_MTRESTE) AS MTRECEP FROM LIGNE WHERE GL_PIECEORIGINE="'+RefOrigine+
             '" AND GL_NATUREPIECEG IN ('+GetPieceAchat+') AND (GL_QTERESTE <> 0 OR GL_MTRESTE <> 0)';

  QQ := OpenSql (requete,true,-1, '', True);
  TRY
    if not QQ.eof then
    begin
      QteRecep  := QQ.findField('QTERECEP').AsFloat;
      MtRecep   := QQ.findField('MTRECEP').AsFloat;
      MtAchat   := QQ.findField('MTACHAT').asFloat;
    end;
  FINALLY
    Ferme (QQ);
  END;

  QteRecep := Qterecep - QteLivre;
  if QteRecep < 0 then QteRecep := 0;
  if QteRecep > 0 then PUA := arrondi(MTachat/QteRecep,V_PGI.OkDecP)
  								else PUA := 0;

  if (QteSais > QteRecep) then
  begin
    PGIBox (TraduireMemoire('Quantité non réceptionnée'),TraduireMemoire('Controle quantitatif'));
    result := false;
  end;

end;

procedure AffecteQtelivrable (TOBOld,TOBL,TOBArticles : TOB;QteInit : double; gestionConso : TGestionPhase);
var QteLivrable : double;
    MtLivrable  : Double;
    TOBArt      : TOB;
begin
  //
  //--- GUINIER ---
  //
  if TOBL.GetValue('GL_TYPELIGNE') <> 'ART' then Exit;
  //
  TOBART := TOBArticles.findFirst (['GA_ARTICLE'], [TOBL.GetValue('GL_ARTICLE')],true);
  if TOBART = nil then Exit;

  if TOBART.GetValue('GA_TYPEARTICLE') = 'PRE' then
  begin
    if TOBART.GetBoolean('GA_RELIQUATMT') then
    begin
      MtLivrable := GetQtelivrable (TOBL,GestionConso,TOBART);
      TOBL.PutValue('GL_MTRELIQUAT', TOBL.GetDouble('GL_MTRESTE') - MtLivrable); // ??
      if TOBL.GetValue('GL_MTRELIQUAT') < 0 then TOBL.PutValue('GL_MTRELIQUAT', 0); // ??
      TOBL.PutValue('GL_MTRESTE', MtLivrable); { NEWPIECE }
      Exit;
    end;
  end;

  QteLivrable := GetQtelivrable (TOBL,GestionConso,TOBART);
  TOBL.PutValue('GL_QTERELIQUAT', TOBL.GetValue('GL_QTERESTE') - QteLivrable); // ??
  if TOBL.GetValue('GL_QTERELIQUAT') < 0 then TOBL.PutValue('GL_QTERELIQUAT', 0); // ??

  TOBL.PutValue('GL_QTEFACT',  QteLivrable);
  TOBL.PutValue('GL_QTESTOCK', QteLivrable); { NEWPIECE }
  TOBL.PutValue('GL_QTERESTE', QteLivrable); { NEWPIECE }

end;

procedure AppliqueDestLivraisonFour (TOBpiece : TOB);
var Indice : integer;
    DestLivrFour : integer;
begin
  DestLivrFour := TOBPiece.GetValue('GP_IDENTIFIANTWOT');
  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBPiece.detail[Indice].Putvalue ('GL_IDENTIFIANTWOL',DestLivrFour);
  end;
end;

function IsModeAvancement (TobPiece : TOB ) : boolean;
var TypeFac: string;
begin
  result := false;
  TypeFac := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
  if ((Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'FBT;FBP')>0) and (TypeFac = 'AVA')) or
    ((Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'DAC;BAC')>0) and (TypeFac = 'DAC')) then
  begin
    result := true;
  end;
end;

function ControleSaisiePieceBtp (TOBPiece : TOB;AFFAIRE : string;TheAction : TactionFiche; transfopiece,duplic : boolean; Silencieux : Boolean = False;IsBordereau : boolean=false) : boolean;

	function IsDevisExistePourAppel(Affaire : string) : boolean;
  var QQ : TQuery;
  begin
  	QQ := OpenSql ('SELECT GP_NUMERO FROM PIECE WHERE GP_NATUREPIECEG="DAP" AND GP_AFFAIRE="'+Affaire+'"',True,1,'',True);
    result := not QQ.eof;
    ferme (QQ);
  end;

var QQ          : Tquery;
		EtatAffaire : String;
    VenteAchat  : String;
    TypeAffaire : String;
    TypeSaisie  : String;
  begin
  result := true;

  //FV1 : 18/06/2014 - FS#921 - DELABOUDINIERE : Revoir les contrôles sur appels et contrats en fonction du code état
  Venteachat := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');

  if (TheAction=TaCreat) or ((Theaction=tamodif) and (Transfopiece or Duplic)) then
  begin
  	if (copy(Affaire,1,1)='W') and (TOBPiece.GetValue('GP_NATUREPIECEG') = 'DAP') then
    begin
      if IsDevisExistePourAppel(Affaire) then
      begin
      	PGIInfo('Impossible : un devis existe déjà pour cet appel');
      	result := false;
        exit;
      end;
    end;
    if (Venteachat = 'VEN') then
    begin
     	if ControleChantiertermineBTP (TOBPiece,Affaire,transfopiece,duplic) then
      begin
        result := false;
        exit;
      end;
    end;
  end;

  //
  if (Affaire='') or (Pos(TOBPiece.GetValue('GP_NATUREPIECEG'),'ETU;DBT;FRC;') <> 0) then exit;

  if (TheAction=TaCreat) then
  Begin
    If (Pos(TOBPiece.GetValue('GP_NATUREPIECEG'), 'FBT;ABT;FBP;ABP;')>0) then TypeSaisie := 'ACH'
    //FV1 - 10/01/2018 - FS#2844 - TEAM - Autoriser la saisie des contre-études sur devis non accepté
    Else if (Pos(TOBPiece.GetValue('GP_NATUREPIECEG'), 'BCE') > 0) then TypeSaisie := 'BCE'
    Else If (Venteachat = 'ACH') then TypeSaisie := Venteachat;
    //
  end;

  //FV1 : 04/04/2018 - FS#3030 - EGCS : Pas de message indiquant "Chantier non accepté" lors de la duplication d'une facture Fournisseur
  if not ControleAffaire(Affaire, 'Saisie de document', TypeSaisie) then
  begin
    result := false;
    exit;
  end;

  {*
  TypeAffaire := copy(UpperCase (Affaire),1,1);


  // verification que l'affaire est bien accepte
  if TypeAffaire='A' then
  begin
   OPTIMISATIONS LS
    QQ := OpenSql ('SELECT AFF_ETATAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="'+Affaire+'"',true);
    if not QQ.eof then
    begin
      if QQ.findField('AFF_ETATAFFAIRE').AsString <> 'ACP' then
      begin
        PgiBox (TraduireMemoire('Ce chantier n''est pas accepté'),TraduireMemoire('Contrôle chantier'));
  			// retour à false uniquement en création pour en pas remettre à blanc
        // le code affaire sur des documents déjà saisis : BRL230107
        if (TheAction=TaCreat) and (Pos(TOBPiece.getValue('GP_NATUREPIECEG'),'CBT;PBT;') > 0) then result := false;
      end;
    end;
  	ferme (QQ);

    EtatAffaire := GetChampsAffaire (Affaire,'AFF_ETATAFFAIRE');

  {  if Pos(EtatAffaire,'ACP;TER')=0 then
    begin
      // retour à false uniquement en création pour en pas remettre à blanc
      // le code affaire sur des documents déjà saisis : BRL230107
      if (TheAction=TaCreat) and (GetParamSocSecur('SO_BTINTERDIREACHATS',False) = True) and (Pos(TOBPiece.getValue('GP_NATUREPIECEG'),GetPieceAchat (true,true,false,true,true)) > 0) then result := false;
      //
      if (TheAction=TaCreat) and (GetParamSocSecur('SO_BTINTERDIREACHATS',False) = True) and (TOBPiece.GetValue('GP_NATUREPIECEG') = 'CBT') then result := false;
      //
      // Contrôle affaire acceptée pour factures et avoirs clients
      if (TheAction=TaCreat) and ((TOBPiece.GetValue('GP_NATUREPIECEG') = 'FBT') or (TOBPiece.GetValue('GP_NATUREPIECEG') = 'ABT')) then result := false;

      if result = false then
        PgiBox (TraduireMemoire('Saisie non autorisée. Ce chantier n''est pas accepté.'),TraduireMemoire('Contrôle chantier'))
      else
        if (not Silencieux) and (not IsBordereau) then PgiBox (TraduireMemoire('Attention, ce chantier n''est pas accepté'),TraduireMemoire('Contrôle chantier'));
    end;
  end;
  *}

  if not result then exit;

  if TOBPiece.getValue ('GP_NATUREPIECEG') = GetparamSoc ('SO_BTNATCHANTIER') then
  begin
    QQ := OpenSql ('SELECT GP_NATUREPIECEG,GP_NUMERO,GP_INDICEG FROM PIECE WHERE GP_AFFAIRE="'+Affaire+
                '" AND GP_NATUREPIECEG="'+TOBPiece.GetValue('GP_NATUREPIECEG')+'"',true,-1, '', True);
    TRY
      if not QQ.eof then
      begin
        if (QQ.FindField ('GP_NATUREPIECEG').asString <> TOBPiece.getValue('GP_NATUREPIECEG')) or
           (QQ.FindField ('GP_NUMERO').asInteger <> TOBPiece.getValue('GP_NUMERO')) or
           (QQ.FindField ('GP_INDICEG').asInteger <> TOBPiece.getValue('GP_INDICEG')) then
        begin
          PgiBox (TraduireMemoire('Une prévision de chantier existe déjà'),TraduireMemoire('Controle prévision de chantier'));
          result := False;
        end;
      end;

    FINALLY
      Ferme (QQ);
    END;
  end;

end;

function ControlQteRecepInAchat (FF : Tform ; TOBL,TOBArt : TOB;Acol,Arow : Integer) : boolean;
var FFAct : TFFacture;
    QteSais,ReliquatInit,OldQTe : Double;
    MtSais,MtReliquatInit,OldMt : Double;
    PUA,coefpapr : double;
begin
  result := True;
  FFAct := TFFacture(FF);
//  if (TOBL.GetValue('GL_PIECEPRECEDENTE') = '') then exit;
  if Pos(TOBL.GetValue('GL_NATUREPIECEG'),GetNaturePieceLBT(false)) <= 0 then exit;

  if TOBL.GetValue('GL_DPA') > 0 then coefpapr := TOBL.GetValue('GL_DPR')/TOBL.GetValue('GL_DPA')
                                 else CoefPaPr := 0;
  With FFact do
  begin
    OldQte := TOBL.GetValue('GL_QTEFACT');
    QteSais := valeur(GS.Cells[Acol,Arow]);
    ReliquatInit := TobL.GetValue('GL_QTERELIQUAT') + OldQte;
    //
    if CtrlOkReliquat(TOBL, 'GL') then
    begin
      OldMT           := TOBL.GetValue('GL_MONTANTHTDEV');
      //MtSais          := valeur(GS.Cells[ACol,Arow]);
      MtReliquatInit  := TobL.GetValue('GL_MTRELIQUAT') + OldMt;
    end;

    //
    if (TOBL.GetValue('GL_TENUESTOCK') <> 'X') then
    begin
      if not TransfoPiece then
      begin
        if ((Action = taModif) and (not DuplicPiece)) then
        begin
          // modification de pièce
//          if not ControleQteBTP (TOBL,TOBL.GetValue('GL_PIECEPRECEDENTE'),QteSais,ReliquatInit,PUA,TaModif) then
					if not ControleQteBTP (TOBL,TOBL.GetValue('GL_PIECEPRECEDENTE'),QteSais,TOBL.GetValue('LIVREORIGINE'),PUA,TaModif) then
          BEGIN
            GS.Cells[ACol, ARow] := StrF00(TOBL.GetValue('GL_QTEFACT'), V_PGI.OkDecQ);
            TOBL.PutValue('GL_RECALCULER', 'X');
            result := false;
            exit;
          END;
          if PUA <> 0 then TOBL.putValue('GL_DPA',PUA);
          if Coefpapr <> 0 then TOBL.Putvalue ('GL_DPR',Arrondi(TOBL.GetValue('GL_DPA')*Coefpapr,V_PGI.okdecP));
        end;
      end else
      begin
        if not DuplicPiece then
        begin
          // transformation de pièce
          if not ControleQteBTP (TOBL,TOBL.GetValue('GL_PIECEPRECEDENTE'),QteSais,ReliquatInit,PUA,TaCreat) then
          BEGIN
            GS.Cells[ACol, ARow] := StrF00(TOBL.GetValue('GL_QTEFACT'), V_PGI.OkDecQ);
            TOBL.PutValue('GL_RECALCULER', 'X');
            result := false;
            exit;
          END;
          if PUA <> 0 then TOBL.putValue('GL_DPA',PUA);
          if Coefpapr <> 0 then TOBL.Putvalue ('GL_DPR',Arrondi(TOBL.GetValue('GL_DPA')*Coefpapr,V_PGI.okdecP));
        end;
      end;
    end;
  end;
end;

procedure ReajustePrixSousDetail (TOBpiece,TOBLF,TOBOuvrage,TOBArticles : TOB; Mode : integer; DEV : RDevise; EnHt : boolean; AvecFrais : boolean=false);
var TOBA,TOBL : TOB;
    Indice,IndPou : integer;
    TypeArticle,ArticleOK : string;
    VenteAchat : string;
    SavPa,SavPr,SAvPv : double;
    Coeffg,CoefMarg,Qte,QteDuDetail,PPQ : double;
    valeurs,ValPou : t_Valeurs;
begin
  ArticleOK:='';
  TOBL:=TOBOuvrage.findfirst(['BLO_TYPEARTICLE'],['POU'],false);
  IndPou := -1;
  if TOBL <> nil then
  begin
  	ArticleOk := TOBL.GetValue('BLO_LIBCOMPL');
    IndPou := TOBL.GetIndex;
  end;

	VenteAchat := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
  for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
    TOBL := TOBOuvrage.detail[Indice];
    //
    TOBL.SetDouble('BLO_COEFFC',0);
    TOBL.SetDouble('BLO_COEFFR',0);
    //
    TypeArticle := TOBL.GetValue ('BLO_TYPEARTICLE');
    Qte :=  TOBL.GetValue ('BLO_QTEFACT');
    Qtedudetail := TOBL.GetValue('BLO_QTEDUDETAIL');
    if TypeArticle = 'POU' Then Continue;
    // suppression de l'article d'écart déjà présent dans le sous-détail
    if TOBL.GetValue ('BLO_CODEARTICLE') = GetParamsoc('SO_BTECARTPMA') then
      begin
      //TOBL.Free;
      continue;
      end;
    if TOBL.detail.count > 0 then
    begin
      ReajustePrixSousDetail (TOBPiece,TOBLF,TOBL,TOBArticles,Mode,DEV,EnHt,AvecFrais);
			CalculOuvFromDetail (TOBL,DEV,valeurs);
      TOBL.Putvalue ('BLO_DPA',Valeurs[0]);
      TOBL.Putvalue ('BLO_DPR',Valeurs[1]);
      TOBL.Putvalue ('BLO_PUHT',Valeurs[2]);
      TOBL.Putvalue ('BLO_PUHTBASE',Valeurs[2]);
      TOBL.Putvalue ('BLO_PUHTDEV',pivottodevise(Valeurs[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBL.Putvalue ('BLO_PUTTC',Valeurs[3]);
      TOBL.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valeurs[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBL.Putvalue ('BLO_PMAP',Valeurs[6]);
      TOBL.Putvalue ('BLO_PMRP',Valeurs[7]);
      //
      TOBL.PutValue('BLO_TPSUNITAIRE',Valeurs[9]);
      TOBL.PutValue('GA_HEURE',Valeurs[9]);
      TOBL.PutValue('BLO_MONTANTFG',Arrondi(Valeurs[13]*Qte,4));
      TOBL.PutValue('BLO_MONTANTFC',Arrondi(Valeurs[15]*Qte,4));
      TOBL.PutValue('BLO_MONTANTFR',Arrondi(Valeurs[14]*Qte,4));
      CalculeLigneAcOuvCumul (TOBL);
      //
      GetValoDetail (TOBL);
      //
    end else
    begin
      // Contrôle si l'article peut être réactualisé ou non selon la fiche article :
      TOBA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBL.GetValue('BLO_ARTICLE')],true);
      if TOBA = nil then Continue;
      if (TOBA.getValue('GA_FANTOME')='X') then continue;

      Savpa := TOBL.getValue('BLO_DPA');
      SavpR := TOBL.getValue('BLO_DPR');
      if EnHt then Savpv := TOBL.getValue('BLO_PUHTDEV')
              else Savpv := TOBL.getValue('BLO_PUTTCDEV');
      if (TOBL.getvalue('BLO_DOMAINE')<>'') and (TOBA.getValue('GA_PRIXPASMODIF')='-') then
      begin
        GetCoefDomaine (TOBL.GetValue('BLO_DOMAINE'),COEFFG,COEFMARG);
      end else
      begin
        Coeffg := TOBA.getValue('GA_COEFFG');
        Coefmarg := TOBA.getValue('GA_COEFCALCHT');
      end;
      if VH_GC.BTCODESPECIF = '001' then
      begin
        CalculeDonneelignePOC (TOBL,coeffg,Coefmarg);
      end;
      //
      InitValoDetail (TOBpiece,TOBL,TOBA,DEV,true);
      TOBA.putValue('DEJA CALCULE','X');
      //
      if ((Mode and 2) = 2) or (AvecFrais) then // Si on reactualise le PR
      begin
         TOBL.putValue('BLO_COEFFG',Coeffg-1);
         TOBL.SetDouble('GL_DPR', Arrondi(TOBL.getValue('BLO_DPA') * Coeffg,V_PGI.okdecP));
      end;
      if ((Mode and 4) = 4) and (TOBLF.getValue('GLC_FROMBORDEREAU')<>'X') then // si on reactualise le PV
      begin
         TOBL.putValue('BLO_COEFMARG',CoefMarg);
         TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('BLO_COEFMARG')-1)*100,2));
      end;
      if (Mode and 1) <> 1 then
      begin
        TOBL.putvalue('BLO_DPA',savpa);
        if TOBL.getvalue('BLO_DPA') <> 0 then TOBL.putValue('BLO_COEFFG',Arrondi((TOBL.getValue('BLO_DPR')/TOBL.getvalue('BLO_DPA'))-1,4))
                                         else TOBL.putValue('BLO_COEFFG',0);
      end;
      if ((Mode and 2) <> 2) and (not AvecFrais)  then
      begin
        TOBL.putvalue('BLO_DPR',savPR);
        if TOBL.getvalue('BLO_DPA') <> 0 then TOBL.putValue('BLO_COEFFG',Arrondi((TOBL.getValue('BLO_DPR')/TOBL.getvalue('BLO_DPA'))-1,4))
                                         else TOBL.putValue('BLO_COEFFG',0);
      end;

      if ((Mode and 4) <> 4) or (TOBLF.getValue('GLC_FROMBORDEREAU')='X') then
      begin
        TOBL.putValue('BLO_COEFMARG',0); // afin de le recalculer
        if EnHt then
        begin
          TOBL.putValue('BLO_PUHTDEV',savpv);
          TOBL.putValue('BLO_PUHT',DEVISETOPIVOTEx(savpv,DEV.Taux,DEV.Quotite,V_PGI.okdecP) );
        end else
        begin
          TOBL.putValue('BLO_PUTTCDEV',savpv);
          TOBL.putValue('BLO_PUTTC',DEVISETOPIVOTEx(savpv,DEV.Taux,DEV.Quotite,V_PGI.okdecP) );
        end;
      end else
      begin
        TOBL.putValue('BLO_COEFMARG',Coefmarg); // afin de le recalculer
      end;
      CalculeLigneAcOuv (TOBL,DEV,true,TOBA);
      if TOBL.GetValue('BLO_PUHT') <> 0 then
      begin
        TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('BLO_PUHT')- TOBL.GetValue('BLO_DPR'))/TOBL.GetValue('BLO_PUHT'))*100,2));
      end else
      begin
        TOBL.PutValue('POURCENTMARQ',0);
      end;
      CalculMontantHtDevLigOuv (TOBL,DEV);
      //
      TOBL.PutValue ('GA_PAHT',TOBL.GetValue('BLO_DPA'));
      TOBL.PutValue ('GA_PVHT',TOBL.GetValue('BLO_PUHT'));
      //
      if TypeArticle = 'PRE' Then TOBL.putvalue('GA_HEURE',TOBL.GetValue('BLO_QTEFACT')) ;
      TOBL.PutValue('GL_RECALCULER','X');
    end;
    //
    if ( IndPou >=0) and (ArticleOKInPOUR (TOBL.GetValue('BLO_TYPEARTICLE'),ArticleOK)) then
    begin
      ValPou[0] := ValPou[0] + Arrondi((Qte/QteDuDetail) * TOBL.GetValue ('BLO_DPA'),V_PGI.okdecP);
      ValPou[1] := ValPou[1] + Arrondi((Qte/QteDuDetail) * TOBL.GetValue ('BLO_DPR'),V_PGI.okdecP);
      ValPou[6] := ValPou[6] + Arrondi((Qte/QteDuDetail) * TOBL.GetValue ('BLO_PMAP'),V_PGI.okdecP);
      ValPou[7] := ValPou[7] + Arrondi((Qte/QteDuDetail) * TOBL.GetValue ('BLO_PMRP'),V_PGI.okdecP);
      if (Mode and 4) = 4 then
      begin
        ValPou[2] := ValPou[2] + Arrondi((Qte/QteDuDetail) * TOBL.GetValue ('BLO_PUHTDEV'),V_PGI.okdecP);
        ValPou[3] := ValPou[3] + Arrondi((Qte/QteDuDetail) * TOBL.GetValue ('BLO_PUTTCDEV'),V_PGI.okdecP);
      end;
    	CalculMontantHtDevLigOuv (TOBL,DEV);
    end;
  end;

  if IndPou >= 0 then
  begin
    PPQ := TOBOuvrage.detail[IndPou].getDouble ('BLO_PRIXPOURQTE');
    Qte := TOBOuvrage.detail[IndPou].GetDouble ('BLO_QTEFACT');
    //
    TOBOuvrage.detail[IndPou].putvalue ('BLO_DPA',ValPou[0]);
    TOBOuvrage.detail[IndPou].putvalue ('BLO_DPR',ValPou[1]);
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUHTDEV',ValPou[2]);
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUTTCDEV',ValPou[3]);
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUHT',DevisetopivotEx(ValPou[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUHTBASE',TOBOuvrage.detail[IndPou].Getvalue ('BLO_PUHT'));
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUTTC',DevisetopivotEx(ValPou[3],DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUTTCBASE',TOBOuvrage.detail[IndPou].Getvalue ('BLO_PUTTC'));
    TOBOuvrage.detail[IndPou].PutValue('BLO_PMAP',ValPou[6]);
    TOBOuvrage.detail[IndPou].PutValue('BLO_PMRP',ValPou[7]);
    CalculMontantHtDevLigOuv (TOBOuvrage.detail[IndPou],DEV);
    //
  end;

end;

function ReajustePrixDevis (TheForm : TForm;TOBTiers,TOBLigneTarif,TOBTarif,TOBBases,TOBBasesL,TOBPiece,TOBArticles,TOBOuvrages,TOBPOrcs: TOB ; DEV : RDevise) : boolean;

	procedure StockeThisLIgne (TOBSauve,TOBL : TOB);
  var TOBS : TOB;
  begin
		TOBL.PutValue('GL_BLOQUETARIF','X');
  	TOBS := TOB.create ('UNE LIGNE',TOBSauve,-1);
    TOBS.AddChampSupValeur ('NOORDRE',TOBL.GEtValue('GL_NUMORDRE'));
  end;

  procedure DepileLouvrage (CalculAuto : boolean ; TOBO : TOB);
  var indice : integer;
  		TOBD : TOB;
  begin
  	for Indice := 0 to TOBO.detail.count -1 do
    begin
    	TOBD := TOBO.detail[Indice];
    	if TOBD.fieldExists('BLO_COEFMARG') and (not CalculAuto) then
      begin
      	TOBD.PutValue('BLO_COEFMARG',0);
      end;
      if TOBD.detail.count > 0 then DepileLouvrage (CalculAuto,TOBD);
    end;
  end;

  procedure DepileLesLignes (CalculAuto : boolean; TOBPiece,TOBOuvrage,TOBSauve : TOB);
  var TOBS,TOBL : TOB;
  		Indice : integer;
      TOBO : TOB;
  begin
    for Indice := 0 to TOBSauve.detail.count -1 do
    begin
    	TOBS := TOBSauve.detail[Indice];
      TOBL := TOBPiece.findFirst(['GL_NUMORDRE'],[TOBS.GetValue('NOORDRE')],true);
      if TOBL <> nil then
      begin
      	if not CalculAuto then
        begin
          TOBL.PutValue('GL_RECALCULER','X');
          TOBL.PutValue('GL_COEFMARG',0);
        end;
      	TOBL.PutValue('GL_BLOQUETARIF','-');
        if TOBL.GetValue('GL_INDICENOMEN')<>0 then
        begin
        	TOBO := TOBOUvrage.detail[TOBL.GetValue('GL_INDICENOMEN')-1];
          if TOBO <> nil then DepileLouvrage (CalculAuto, TOBO);
        end;
      end;
    end;
  end;


var FF : TFFacture;
    i : integer;
    TOBL, TOBA,TOBO, TOBSauve: TOB;
    TypeArticle,VenteAchat,RecupPrix : string;
    Valeurs : T_Valeurs;
//    SavPrix : double;
    EnHt,EnPa : boolean;
    Mode : integer;
    SavPa,SavPr,SAvPv,lastPv,NewPv : double;
    Coeffg,CoefMarg,Qte : double;
    ExistFr,ExistFC : boolean;
begin
  VenteAchat := GetInfoParPiece(TOBpiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
  ExistFR := IsExisteFraisRepartit (TOBPorcs);
  ExistFC := (TOBPiece.getValue ('GP_PIECEFRAIS')<>'');
	TOBSauve := TOB.Create ('LES LIGNES SAUV',nil,-1);
	result := false;
  if TOBPiece.Detail.Count <= 0 then Exit;
  if TOBArticles.Detail.Count <= 0 then Exit;
  FF := TFFacture(TheForm);

  EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');

  RecupPrix := GetInfoParPiece(TOBPIece.GetValue('GP_NATUREPIECEG'), 'GPP_APPELPRIX');
  EnPa := ((RecupPrix='PAS') or (RecupPrix = 'DPA'));
  DefiniModeReactualisation (Mode,TOBPiece.GetValue('_BLOQUETARIF'),TOBPiece);
  if Mode > 0 then
  begin
  	result := true;
    //
    TOBPiece.setDouble('GP_COEFFC',0);
    TOBPiece.setDouble('GP_COEFFR',0);
    //
    for i := 0 to TOBPiece.Detail.Count - 1 do
    begin
      if ((FF.GS.NbSelected = 0) or (FF.GS.IsSelected(i + 1))) then
      begin
        TOBL := TOBPiece.detail[i];
        //
        TOBL.setDouble('GL_COEFFC',0);
        TOBL.setDouble('GL_COEFFR',0);
        //
        if (TOBL.GetValue('GL_TYPELIGNE') <> 'ART') and (TOBL.GetValue('GL_TYPELIGNE') <> 'ARV') Then Continue;
        TypeArticle := TOBL.GetValue ('GL_TYPEARTICLE');
        if TypeArticle = 'POU' Then continue;
        //
        if ((Mode and 4) = 4) and (TOBL.getValue('GLC_FROMBORDEREAU')<>'X') and (TOBL.getValue('GL_BLOQUETARIF')<>'X' ) then // si on reactualise le PV
        begin
          FF.deduitLaLigne (TOBL);
          FF.InitMontantLigne (TOBL);
        end;
        //
        if (TypeArticle = 'OUV') or (TypeArticle = 'ARP') then
        begin
          TOBO := TOBOuvrages.detail[TOBL.GetValue('GL_INDICENOMEN') -1];
          if TOBO.detail.count =  0 then continue;
          ReajustePrixSousDetail (TOBPiece,TOBL,TOBO,TOBArticles,Mode,DEV,EnHt,(ExistFR or ExistFC));
          InitTableau (Valeurs);
          CalculeOuvrageDoc (TOBO,1,1,true,DEV,valeurs,EnHt);
      		GetValoDetail (TOBO); // pour le cas des Article en prix posés
          Qte := TOBL.Getvalue('GL_QTEFACT');
          TOBL.Putvalue('GL_MONTANTPAFG',valeurs[10]*Qte);
          TOBL.Putvalue('GL_MONTANTPAFR',valeurs[11]*Qte);
          TOBL.Putvalue('GL_MONTANTPAFC',valeurs[12]*Qte);
          TOBL.Putvalue('GL_MONTANTFG',valeurs[13]*Qte);
          TOBL.Putvalue('GL_MONTANTFR',valeurs[14]*Qte);
          TOBL.Putvalue('GL_MONTANTFC',valeurs[15]*Qte);
          TOBL.Putvalue('GL_MONTANTPA',Arrondi((Qte * TOBL.GetValue('GL_DPA')),V_PGI.okdecV));
          TOBL.Putvalue('GL_MONTANTPR',Arrondi((Qte * TOBL.GetValue('GL_DPR')),V_PGI.okdecV));

          TOBL.Putvalue('GL_DPA',valeurs[16]);
          TOBL.Putvalue('GL_DPR',valeurs[17]);
          //
          TOBL.Putvalue('GL_DPA',valeurs[0]);
          TOBL.Putvalue('GL_DPR',valeurs[1]);
          TOBL.Putvalue('GL_PMAP',valeurs[6]);
          TOBL.Putvalue('GL_PMRP',valeurs[7]);
          TOBL.putvalue('GL_TPSUNITAIRE',valeurs[9]);
          TOBL.putvalue('GL_RECALCULER', 'X');
          if ((Mode and 4) = 4) and (TOBL.getValue('GLC_FROMBORDEREAU')<>'X') and (TOBL.getValue('GL_BLOQUETARIF')<>'X' ) then
          begin
            TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
            TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
            TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
            TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBL.GetValue('GL_PUTTCDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
            TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
            TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
						TOBL.putValue('GL_COEFMARG',0);
          end;
          StockeInfoTypeLigne (TOBL,Valeurs);
        end else
        begin
          // Contrôle si l'article peut être réactualisé ou non selon la fiche article :
          TOBA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBL.GetValue('GL_ARTICLE')],true);
          if TOBA = nil then Continue;
          if (TOBA.getValue('GA_FANTOME')='X') then continue;

          Savpa := TOBL.getValue('GL_DPA');
          SavpR := TOBL.getValue('GL_DPR');
          //
          if EnHt then Savpv := TOBL.getValue('GL_PUHTDEV')
                  else Savpv := TOBL.getValue('GL_PUTTCDEV');
          //
          if (TOBL.getvalue('GL_DOMAINE')<>'') and (TOBA.getValue('GA_PRIXPASMODIF')='-') then
          begin
						GetCoefDomaine (TOBL.GetValue('GL_DOMAINE'),COEFFG,COEFMARG);
          end else
          begin
          	Coeffg := TOBA.getValue('GA_COEFFG');
          	Coefmarg := TOBA.getValue('GA_COEFCALCHT');
          end;
          //
          if VH_GC.BTCODESPECIF = '001' then
          begin
            CalculeDonneelignePOC (TOBL,COEFFG,COEFMARG);
          end;
          //
          if (Mode and 1) = 1 then // Si on reactualise le PA
          begin
             TOBL.putValue('GL_DPA',0);
          end;
          if ((Mode and 2) = 2) or (ExistFR) or (ExistFR) then // Si on reactualise le PR
          begin
             TOBL.putValue('GL_COEFFG',Coeffg-1);
          end;
          //
          if ((Mode and 4) = 4) and (TOBL.getValue('GLC_FROMBORDEREAU')<>'X') and (TOBL.getValue('GL_BLOQUETARIF')<>'X' ) then
          begin
            if CoefMarg <> 0 then
            begin
            	TOBL.putValue('GL_COEFMARG',CoefMarg);
              TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
            end else
            begin
							TOBL.putValue('GL_PUHT',TOBA.GetDouble('GA_PVHT'));
       				TOBL.PutValue('GL_PUHTDEV',PivotToDevise(TOBL.GetValue('GL_PUHT'),DEV.Taux,DEV.Quotite,DEV.Decimale)) ;
              if TOBL.getValue('GL_DPR') <> 0 then
              begin
								TOBL.putValue('GL_COEFMARG',Arrondi(TOBL.getValue('GL_PUHT')/TOBL.getValue('GL_DPR'),4));
                TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
              end else
              begin
								TOBL.putValue('GL_COEFMARG',0);
              end;
            end;
          end;
          //
					InitValoArtNomen (TOBA);
          TOBA.putValue('DEJA CALCULE','X');
					AffectePrixValo (TOBL,TOBA,nil);
          //
          if (Mode and 1) <> 1 then
          begin
          	TOBL.putvalue('GL_DPA',savpa);
            if TOBL.getvalue('GL_DPA') <> 0 then
            	TOBL.putValue('GL_COEFFG',arrondi((TOBL.getValue('GL_DPR')/TOBL.getvalue('GL_DPA'))-1,4));
          end;
          if (Mode and 2) <> 2 then
          begin
          	TOBL.putvalue('GL_DPR',savPR);
            if TOBL.getvalue('GL_DPA') <> 0 then
            	TOBL.putValue('GL_COEFFG',arrondi((TOBL.getValue('GL_DPR')/TOBL.getvalue('GL_DPA'))-1,4));
          end;
          //
          if (ExistFr) or (ExistFC) then
          begin
             TOBL.putValue('GL_COEFFG',Coeffg-1);
             TOBL.putValue('GL_COEFMARG',Coefmarg);
          end;
          //
          CalculeLigneAc (TOBL,DEV);
          //
          if ((Mode and 4) = 4) and (TOBL.getValue('GLC_FROMBORDEREAU')<>'X') and (TOBL.getValue('GL_BLOQUETARIF')<>'X' ) then
          begin
          	if EnHt then LastPv := TOBL.GetValue('GL_PUHTDEV')
            				else LastPv := TOBL.GetValue('GL_PUTTCDEV');
            AppliquePrixFromTarif (FF,TOBPiece,TOBL,TOBBases,TOBBasesL,TOBouvrages,EnPa,EnHT,Savpv,DEV,i+1);
          	if EnHt then newPv := TOBL.GetValue('GL_PUHTDEV')
            				else NewPv := TOBL.GetValue('GL_PUTTCDEV');
            if LastPv <> newPv then TOBL.putValue('GL_COEFMARG',0); // afin de le recalculer
          end else
          begin
          	TOBL.putValue('GL_COEFMARG',0); // afin de le recalculer
          	if EnHt then
            begin
              TOBL.putValue('GL_PUHTDEV',savpv);
              TOBL.putValue('GL_PUHT',DEVISETOPIVOTEx(savpv,DEV.Taux,DEV.Quotite,V_PGI.okdecP) );
            end else
            begin
              TOBL.putValue('GL_PUTTCDEV',savpv);
              TOBL.putValue('GL_PUTTC',DEVISETOPIVOTEx(savpv,DEV.Taux,DEV.Quotite,V_PGI.okdecP) );
            end;
          end;
          
          if TOBL.GetValue('GL_PUHT') <> 0 then
          begin
            TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
          end else
          begin
            TOBL.PutValue('POURCENTMARQ',0);
          end;
          //
          TOBL.putvalue('GL_RECALCULER', 'X');
        end;
      end;
    end;
    FF.GS.ClearSelected;

    TOBL := TOBPiece.findFirst(['GL_CODEARTICLE'],[GetParamsoc('SO_BTECARTPMA')],True);
    if TOBL<> nil then
    begin
      FF.BTPClickDel (TOBL.GetValue('GL_NUMLIGNE'));
      TOBL.free;
    end;
  	FF.ReinitPieceForCalc ((not (ExistFR or ExistFC)) and ((Mode and 4) <> 4));
  end;
	TOBSauve.free;
end;

function GetQteResteCdeOrigine (TOBL : TOB) : double;
var Cledoc : R_CLEDOC;
    QQ : Tquery;
    Req : String;
begin
  result := 0;
  if TOBL.GetValue('GL_PIECEPRECEDENTE')='' then exit;
  DecodeRefPiece (TOBL.GetValue('GL_PIECEPRECEDENTE'),cledoc);
  req := 'SELECT GL_QTERESTE FROM LIGNE WHERE '+WherePiece(cledoc,ttdLigne,true);
  QQ := OpenSql (Req,true,-1, '', True);
  if not QQ.eof then result := QQ.findField('GL_QTERESTE').AsFloat;
  ferme (QQ);
end;

function GetMtResteCdeOrigine (TOBL : TOB) : double;
var Cledoc : R_CLEDOC;
    QQ : Tquery;
    Req : String;
begin
  result := 0;

  if TOBL.GetValue('GL_PIECEPRECEDENTE')='' then exit;

  DecodeRefPiece (TOBL.GetValue('GL_PIECEPRECEDENTE'),cledoc);

  req := 'SELECT GL_MTRESTE FROM LIGNE WHERE '+WherePiece(cledoc,ttdLigne,true);

  QQ := OpenSql (Req,true,-1, '', True);

  if not QQ.eof then result := QQ.findField('GL_MTRESTE').AsFloat;

  ferme (QQ);

end;


procedure AjouteLesLignesBTP (FF: Tform ;TOBpiece,TOBAffaire,TobTiers,TOBArticles: TOB;TheDoc : T_VisuDoc; GestionConso : TGestionPhase);
var Indice : integer;
		TOBL,TOBR,TOBPrec,TOBA,TOBNA : TOB;
    PositIns,NiveauImbric : integer;
begin
	For indice := 0 to TheDOC.TOBresult.detail.count -1 do
  begin
    PositIns := TOBPiece.detail.count;
    if PositIns > 0 then TOBPrec := TOBPiece.detail[PositIns-1];
    //
    TOBR := TheDoc.TOBresult.detail[Indice];
    TOBA := TheDoc.TOBArticles.FindFirst (['GA_ARTICLE'],[TOBR.GEtVAlue('GL_ARTICLE')],true);
    if TOBR = nil then break;
    if PositIns = 0 then
    begin
    	NiveauImbric := 0;
    end else
    begin
    	if TOBprec <> nil then
      begin
        if IsFinParagraphe (TOBPrec) then NiveauImbric := TOBPrec.GetValue('GL_NIVEAUIMBRIC')-1
                                     else NiveauImbric := TOBPrec.GetValue('GL_NIVEAUIMBRIC');
      end else NiveauImbric := 0;
    end;
    InsertTobLigne (TOBPiece,0);
    TOBL := GetTOBLIgne (TOBPiece,TOBPiece.detail.count);
    TOBL.dupliquer (TOBR,false,true);
    AddLesSupLigne(TOBL,False) ;  InitLesSupLigne (TOBL);
    NewTOBLigneFille (TOBL);
    PieceVersLigne (TOBPIece,TOBL);
    AffaireVersLigne (TOBPiece,TOBL,TOBAffaire);
    TOBL.PutValue('GL_NIVEAUIMBRIC',NiveauImbric);
    TOBL.putValue('GL_RECALCULER','X');
    TOBL.PutValue('GL_NUMORDRE',GetMaxNumOrdre(TOBPiece));
    TOBL.PutValue('GL_NUMLIGNE',TOBPiece.detail.count);
    // ici on a récupérer le PU Achat en provenance de la réception
    // Il faut donc recalculer le PUV avec les coefs de vente.
    if TOBL.GetValue('GL_PUHTDEV') = 0 then TOBL.PutValue('GL_PUHTDEV',TOBL.GetValue('GL_DPR'));
    if TOBA <> nil then
    begin
    	// on réapplique les coefs a partie du pa
      if TOBA.GEtValue('GA_COEFFG')<> 0 then TOBL.PutValue('GL_DPR',Arrondi(TOBL.GetValue('GL_DPA') * TOBA.GetValue('GA_COEFFG'),V_PGI.OkdecP));
    	if TOBA.GEtVAlue('GA_COEFCALCHT') <> 0 then TOBL.putValue('GL_PUHTDEV',Arrondi(TOBL.GetValue('GL_DPR') * TOBA.GetValue('GA_COEFCALCHT'),V_PGI.OkdecP));
    end;
    // Insertion des affectations de receptions dans les consos
    gestionConso.DefiniReceptionsFromHlienAssocie (TOBL,TOBR);
    TOBL.PutValue ('DIRECTFOU',gestionConso.GetQteLivrable(TOBL));
    // --
    if TOBA <> nil then
    begin
      TOBNA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBA.GetValue('GA_ARTICLE')],true);
      if TOBNA = nil then
      begin
      	TOBNA := TOB.Create ('ARTICLE',TOBArticles,-1);
        TOBNA.dupliquer (TOBA,True,true);
				DispoChampSupp (TOBNA);
      end;
    end;
  end;
  TOBPiece.putValue('GP_RECALCULER','X');
  if FF is TFFActure then TFFacture(FF).CalculeLaSaisie (-1,-1,true);
end;

procedure PositionneModeBordereauBTP (FF: Tform; TOBPiece : TOB);
var FFact : TFFacture;
begin
  FFact := TFFacture (FF);
  with FFAct do
  begin
    if TOBpiece.detail.count > 0 then
    begin
      if (TOBPiece.detail[0].getValue('GL_BLOQUETARIF')='X') and (TOBPiece.detail[0].getValue('GLC_FROMBORDEREAU')<>'X') then
      begin
        TOBPiece.PutValue('_BLOQUETARIF','X');
      end;
    end;
    if TOBPiece.GetValue('_BLOQUETARIF')='-' Then
    begin
      TmenuItem(FindComponent('Modebordereau')).caption := 'Mode P.V bloqué';
      TToolBarButton97(FindComponent('BprixMarche')).enabled := true;
      caption := TitreInitial;
      refresh;
    end else
    begin
      if (not FFact.SaisieTypeAvanc) then TToolBarButton97(FindComponent('BprixMarche')).enabled := false;
      TmenuItem(FindComponent('Modebordereau')).caption := 'Mode standard';
      caption := TitreInitial + ' (Mode Bordereau)';
      refresh;
    end;
  end;
end;

procedure ClickValideAndStayHere (FF : Tform ;TOBPiece, TOBNomenclature, TobOuvrage,TOBTiers, TOBAffaire,TOBLigneRg : TOB);
var X : TFFacture;
//		Splash : TFsplashScreen;
		ReinitQte : boolean;
begin
	if not (FF is TFFacture) then exit;

	X := TFFacture(FF);
  if X.InValidation then exit;

  X.Invalidation := true;
  TRY
    X.Label1.Caption  := TraduireMemoire('Validation en cours...');
    X.Animate1.Active := true;
    X.WAIT.Visible := true;
    X.WAIT.Refresh;
    X.Enabled := false;

    if (X.TransfoMultiple) or (X.TransfoPiece) or (X.DuplicPiece) or (X.SaisieTypeAvanc) then
    begin
       X.ClickValide;
       Exit;
    End;
    //
    X.ClickValide(True);
    if X.modeAudit then X.AuditPerf.Debut('Suppression des sous détails');
    X.SupLesLibDetail(TOBPiece);
    if X.modeAudit then X.AuditPerf.Fin('Suppression des sous détails');
    if X.ModifAvanc then ReinitQte := true else ReinitQte := false;
    AfterImprimePieceBtp (FF,TOBPiece,TOBNomenclature,TOBOuvrage,TOBTiers,TOBAffaire,TOBLIgneRG,ReinitQte);
    X.Bvalider.enabled := True;
  FINALLY
    X.enabled := true;
    X.Animate1.Active := false;
    X.WAIT.Visible := false;
    X.Invalidation := false;
  END;
end;

procedure InsereLigneRefPrecedent (TOBPiece : TOB ;OldCledoc : R_Cledoc);
var TOBL : TOB;
    maxNumOrdre : integer;
begin
	G_LigneCommentaireBTP (TOBPiece,OldCledoc);
  TOBL := TOBPiece.detail[0];
  if TOBL.getInteger('GL_NUMORDRE')=0 then
  begin
    MaxNumOrdre:=LireMaxNumOrdre(TOBPiece);
    Inc(MaxNumOrdre);
    TOBL.PutValue('GL_NUMORDRE', MaxNumOrdre);
    EcrireMaxNumordre (TOBPiece,MaxNumOrdre);
  end;
end;

procedure AppliquePrixFromTarif (FF : TForm;TOBPiece,TOBL,TOBBases,TOBBasesL,TOBouvrage : TOB;EnPa,EnHT : boolean;Savprix : double;DEV : RDevise; ARow : integer);
var XX : TFFacture;
begin
	if SavPrix = 0 then exit;
  //
	XX := TFFacture(FF);
	if (TOBL.GetValue('FROMTARIF')='X') or ((tModeSaisieBordereau in XX.SaContexte) and (XX.OrigineEXCEL))    then
  begin
    if XX.GP_FACTUREHT.Checked then
    begin
      {$IFDEF BTP}
      if ((TOBL.getValue('GL_TYPEARTICLE') = 'OUV') or (TOBL.getValue('GL_TYPEARTICLE') = 'ARP')) and
        ApplicPrixPose (TOBpiece) then
      begin
        TraitePrixOuvrage(TOBPiece,TOBL,TOBBases,TOBBasesL,TOBOuvrage, EnHt, SavPrix,0,DEV,EnPA);
        if SG_PX >= 0 then
        begin
          if EnPa then
          begin
            XX.GS.cells[SG_px, Arow] := strf00(TOBL.GetValue('GL_DPA'), V_PGI.OkdecP);
          end else
          begin
            XX.GS.cells[SG_px, Arow] := strf00(TOBL.GetValue('GL_PUHTDEV'), V_PGI.OkdecP);
          end;
        end;
      end else
      {$ENDIF}
        TOBL.PutValue('GL_PUHTDEV', SavPrix);
    end else
    begin
      {$IFDEF BTP}
      if ((TOBL.getValue('GL_TYPEARTICLE') = 'OUV') or (TOBL.getValue('GL_TYPEARTICLE') = 'ARP')) and
        ApplicPrixPose (TOBpiece) then
      begin
        TraitePrixOuvrage(TOBPiece,TOBL,TOBBases,TOBBasesL,TOBOuvrage, EnHT, SavPrix,0,DEV,EnPa);
        if SG_PX >= 0 then
        begin
          if EnPa then
          begin
            XX.GS.cells[SG_px, Arow] := strf00(TOBL.GetValue('GL_DPA'), V_PGI.OkdecP);
          end else
          begin
            XX.GS.cells[SG_px, Arow] := strf00(TOBL.GetValue('GL_PUTTCDEV'), V_PGI.OkdecP);
          end;
        end;
      end else
      {$ENDIF}
        TOBL.PutValue('GL_PUTTCDEV', SavPrix);
    end;
  	TOBL.PutValue('FROMTARIF','-');
  	TOBL.PutValue('GL_RECALCULER','X');
  end;
end;

procedure  WarningSiPrevisionnel (FF : Tform;TOBpiece,TOBLienDevCha :TOB);
var XX : TFFacture;
		TheMessage : String;
begin
	XX := TFFacture(FF);
  if ((TOBPiece.getValue('GP_NATUREPIECEG')=VH_GC.AFNatAffaire) OR (TOBPiece.getValue('GP_NATUREPIECEG')='BCE')) and
  	 (TOBLienDevCha.detail.count > 0) then
  begin
  	if (XX.DuplicPiece) or (XX.SaisieTypeAvanc) then exit;
//    TheMEssage := TraduireMemoire('ATTENTION : Une prévision de chantier (ou une contre étude) a déjà été générée pour ce document.#13#10'+
//    															'Toute modification de ce document entrainera la perte du lien avec celle-ci');
    TheMEssage := TraduireMemoire('ATTENTION : Une prévision de chantier (ou une contre étude) a déjà été générée pour ce document.#13#10'+
    															'Vous devrez mettre à jour manuellement la prévision de chantier (ou la contre étude)');
  	PGIInfo(TheMessage,XX.Caption);
  end;
end;

procedure	WarningFinSiprevisionnel (FF : Tform;TOBpiece,TOBLienDevCha :TOB);
var XX : TFFacture;
		TheMessage : String;
begin
	XX := TFFacture(FF);
  if (not XX.SaisieTypeAvanc ) and (TOBPiece.getValue('GP_NATUREPIECEG')=VH_GC.AFNatAffaire) and (TOBLienDevCha.detail.count > 0) then
  begin
//    TheMEssage := TraduireMemoire('ATTENTION : la modification de ce document entraine la perte du lien avec la prévision de chantier (ou de la contre étude.#13#10'+
//    															'Veuillez mettre à jour manuellement la prévision de chantier (ou la contre étude)');
    TheMEssage := TraduireMemoire('ATTENTION : la modification de ce document ne met pas à jour automatiquement la prévision de chantier (ou de la contre étude.#13#10'+
    															'Veuillez mettre à jour manuellement la prévision de chantier (ou la contre étude)');
  	PGIInfo(TheMessage,XX.Caption);
    (*
    TOBlienDevCha.clearDetail;
    ExecuteSql('DELETE FROM LIENDEVCHA WHERE BDA_REFD LIKE "'+EncodeLienDevCHADel(TOBPiece)+'"');
    *)
  end;
end;

function IsPieceOrigineEtude (TOBL : TOB) : boolean;
var Cledoc : r_cledoc;
		RefPiece : string;
begin
	result := false;
  refPiece := TOBL.getValue('GL_PIECEORIGINE');
  if refpiece = '' then exit;
	DecodeRefPiece (TOBL.getValue('GL_PIECEORIGINE'),cledoc);
  result := (Cledoc.NaturePiece='ETU');
end;

procedure ReinitCoefPaPr(TOBL,TOBouvrage : TOB);
begin
(*
  TOBL.PutValue('GL_COEFFG',0);
  TOBL.PutValue('GL_COEFFC',0);
  TOBL.PutValue('GL_DPR',TOBL.GetValue('GL_DPA'));
  if TOBL.GetValue('GLC_FROMBORDEREAU')<>'X' then
  begin
    TOBL.PutValue('GL_PUHTDEV',TOBL.GetValue('GL_DPA'));
    TOBL.PutValue('GL_PUTTCDEV',TOBL.GetValue('GL_DPA'));
  end;
  if TOBL.GetValue('GL_BLOQUETARIF')<>'X' then
  begin
    TOBL.PutValue('GL_PUHTDEV',TOBL.GetValue('GL_DPA')*TOBL.GetValue('GL_COEFMARG'));
  end;
  if IsOuvrage (TOBL) then ReinitCoefPaPrOuv(TOBL,TOBouvrage);
*)
end;

procedure ReactiveEtude(TOBPiece : TOB);
var result : boolean;
		Requete,AffaireDevis : string;
    cledoc : r_cledoc;
    Indice : integer;
    TOBL,TOBRef,TOBP : TOB;
    QQ : TQuery;
begin

  TOBRef := nil;
	if TOBPiece.getValue('GP_NATUREPIECEG')='DBT' Then    // cas de la supression d'un devis
  begin
  	//
    for Indice := 0 to TOBPiece.detail.count -1 do
    begin
    	TOBL := TOBPiece.detail[Indice];
      if TOBL.GetValue('GL_PIECEPRECEDENTE')<> '' then BEGIN TOBREF := TOBL;break; END;
    end;

    if TOBREF <> NIL then
    begin
  		DecoderefPiece (TOBRef.getValue('GL_PIECEPRECEDENTE'),cledoc);

      requete := 'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+WherePiece(Cledoc,ttdpiece,true);
      QQ := OpenSql (requete,True,-1, '', True);
      AffaireDevis := QQ.findField('GP_AFFAIREDEVIS').asString;
      ferme (QQ);

      requete := 'UPDATE PIECE SET GP_VIVANTE="X" WHERE '+WherePiece (cledoc,ttdPiece,true);

      result := (ExecuteSql (requete)>0);
      if (not result) then
      begin
        V_PGI.IOError := OeUnknown;
        exit;
      end;
      //
      requete :=  'UPDATE LIGNE SET GL_VIVANTE="X" WHERE '+WherePiece(cledoc,ttdligne,false,false);
      result := (ExecuteSql (requete)>0);
      if (not result) then
      begin
        V_PGI.IOError := OeUnknown;
        exit;
      end;
      //
      Requete := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE="ENC" WHERE AFF_AFFAIRE="'+Affairedevis+'"';
      result := (ExecuteSql (requete)>0);
      if (not result) then
      begin
        V_PGI.IOError := OeUnknown;
        exit;
      end;
    end;
  end;
end;

function ErreurSiZero (TOBPiece : TOB) : boolean;
begin
	result := true;
	if (Pos(TOBPiece.getValue('GP_NATUREPIECEG'),PieceFacturation) > 0) and (TOBPiece.getValue('GP_TOTALHT')=0) then
  begin
    if PgiAsk('ATTENTION : La pièce est d''un montant nul#13#10Valider quand même ?')<> mryes then
    begin
			result := false;
    end;
  end;
end;

procedure ControleEtReajustePieceOrigine (TOBPiece : TOB; TOBArticles : TOB);
var Indice : Integer;
		TOBL,TOBLO,TOBA,TOBDispo : TOB;
    cledoc : r_cledoc;
    QQ,TQDepot : TQuery;
    colPlus,colMoins,depot,NewDepot,Article : string;
    Qte,RatioVa : double;
begin
	TOBLO := TOB.create ('LIGNE',nil,-1);
	For Indice := 0 to TOBPiece.detail.count -1 do
  begin
  	TOBLO.ClearDetail;
  	TOBL := TOBPIece.detail[Indice];
    if (TOBL.GetVAlue('GL_TENUESTOCK')='-') or (TOBL.GetValue('GL_ARTICLE')='') then continue;
    if TOBL.GetValue('GL_PIECEORIGINE') = '' then continue;
    DecodeRefPiece (TOBL.GetValue('GL_PIECEORIGINE'),cledoc);
    if GetInfoParPiece (cledoc.NaturePiece,'GPP_VENTEACHAT') <> 'VEN' then continue;
    QQ := OpenSql ('SELECT * FROM LIGNE WHERE '+WherePiece(cledoc,ttdligne,true,true),true,-1, '', True);
    if not QQ.eof then
    begin
    	TOBLO.selectDb ('',QQ);
      // changement de depot en cours de route ...grrrr
      ColPlus := GetInfoParPiece(cledoc.NaturePiece, 'GPP_QTEPLUS');
      ColMoins := GetInfoParPiece(cledoc.NaturePiece, 'GPP_QTEMOINS');
      QTe := TOBLO.GetValue('GL_QTERESTE');
      Depot := TOBLO.getValue('GL_DEPOT');
      NewDepot := TOBL.GetValue('GL_DEPOT');
      Article := TOBL.GetValue('GL_ARTICLE');
      if (Depot <> NewDepot) and (Qte>0) then
      begin
        TOBA:=TOBArticles.FindFirst(['GA_ARTICLE'],[Article],False) ;
        TOBdispo := TOBA.findFirst(['GQ_ARTICLE','GQ_DEPOT'],[Article,Depot],true);
        if TOBDispo <> nil then
        begin
        	RatioVA := GetRatio(TOBLO, nil, trsStock);
        	UpdateQteStock(TOBDispo,Qte,RatioVA,colPlus,colMoins,false) ;
        end;
        TOBLO.putValue('GL_DEPOT',NewDepot);
        TOBdispo := TOBA.findFirst(['GQ_ARTICLE','GQ_DEPOT'],[TOBLO.GEtVAlue('GL_ARTICLE'),newDepot],true);
        if TOBDispo = nil then
        begin
        	TOBDisPO := TOB.create ('DISPO',TOBA,-1);
        	TQDepot := OpenSQL('SELECT * FROM DISPO WHERE GQ_ARTICLE ="'+Article+ '" AND ' +
      											 'GQ_DEPOT="' + NewDepot + '" AND GQ_CLOTURE="-"', True,-1, '', True);
          if not TQDepot.EOF then
          begin
          	TOBDispo.LoadDetailDB('DISPO', '', '', TQDepot, True, True);
          end else
          begin
          	TOBDispo.putValue('GQ_ARTICLE',Article);
          	TOBDispo.putValue('GQ_DEPOT',NewDepot);
          	TOBDispo.putValue('GQ_DPA',TOBL.GetValue('GL_DPA'));
          	TOBDispo.putValue('GQ_PMAP',TOBL.GetValue('GL_PMAP'));
          	TOBDispo.putValue('GQ_DPR',TOBL.GetValue('GL_DPR'));
          	TOBDispo.putValue('GQ_PMRP',TOBL.GetValue('GL_PMRP'));
          	TOBDispo.putValue('GQ_DATECREATION',V_PGI.DateEntree);
          end;
          Ferme(TQDepot);
        end;
        UpdateQteStock(TOBDispo,Qte,RatioVA,colPlus,colMoins,true) ;
        TOBLO.PutValue('GL_DEPOT',NewDepot);
        TRY
          TOBLO.UpdateDB;
        EXCEPT
          on E: Exception do
          begin
            PgiError('Erreur Mise a jour : ' + E.Message, 'Reajustement stock Origine');
          end;
        END;
      end;
    end;
    ferme (QQ);
  end;
  TOBLO.free;
end;

function IsModifiable (TOBPiece : TOB) : boolean;

function LigneInCommandeFou (TOBL : TOB) : boolean;
var QQ : TQuery;
		RefPiece : string;
begin
	result := false;
	if V_PGI.Superviseur then exit;
	if TOBL.GetValue('GL_AFFAIRE') = '' then exit;
  RefPiece := EncodeRefPiece (TOBL);
  QQ := OPenSql ('SELECT GL_NUMORDRE FROM LIGNE WHERE GL_NATUREPIECEG="CF" AND '+
  							 'GL_AFFAIRE="'+TOBL.GetValue('GL_AFFAIRE')+'" AND GL_PIECEORIGINE="'+RefPiece+'"',True,-1, '', True);
  result := not QQ.eof;
  ferme (QQ);
  if result then exit;
  QQ := OPenSql ('SELECT BAD_REFGESCOM FROM DECISIONACHLIG LEFT JOIN DECISIONACH ON BAE_NUMERO=BAD_NUMERO WHERE BAD_REFGESCOM="'+RefPiece+'" AND BAE_VALIDE="-"',True,-1, '', True);
  result := not QQ.eof;
  ferme (QQ);
end;

var indice : integer;
		TOBL : TOB;
begin
	result := true;
  
	if (TOBPiece.getValue('GP_NATUREPIECEG')='DBT') and GetParamSocSecur('SO_BTBLOQUEDEVISACCEPTE', False) and (TOBPiece.GetValue('ETATDOC')='ACP') then
  begin
    Result := False;
  end;

	if TOBPiece.getValue('GP_NATUREPIECEG')<>'CBT' then exit;
	if TOBPiece.getValue('GP_AFFAIRE')='' then exit;
  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
  	TOBL := TOBPiece.detail[Indice];
    if TOBL.getValue('GL_TYPELIGNE') <> 'ART' Then continue;
    if TOBL.GetValue('GL_ARTICLE')='' then continue;
    if LigneInCommandeFou (TOBL) then
    begin
    	result := false;
      break;
    end;
  end;
end;

procedure MemoriseListeSaisieBTP (FF : TForm);
//var FFact : TFFacture;
//		TheColonnes : string;
begin
(*
  FFAct := TFFacture (FF);
  with FFact do
  begin
  	LesColonnesVte := GS.titres[0];
  	LesColonnesAch := GS.titres[0];
  end;
*)
end;

procedure ReinitListeSaisieBTP (FF : TForm);
var FFact : TFFacture;
begin
  FFAct := TFFacture (FF);
  with FFact do
  begin
		GS.titres[0] := LesColonnesVte;
  end;
end;

function PieceUniteautorisee (Nature : string) : boolean;
begin
  result := (Pos(Nature,'ETU;DBT;FBT;ABT;ABP;AFF;BCE;FPR;FAC;DEF;CF;BLF;FF;CFR;LFR;DE;CC;BLC;FBC')>0);
end;

function ControleMargeBTP (TOBPiece : TOB; AfficheMessage : Boolean=False) : boolean;
Var CftP, CftC : Double;
		Mess : String;

begin
	result := true;

	if TOBPiece.GetValue('GP_VENTEACHAT') <> 'VEN' then Exit;

  CftP:=GetparamsocSecur('SO_BTMARGEMINI',0.0);
	if CftP=0.0 then exit;
  if TobPiece.Getvalue('GP_MONTANTPA')=0.0 then
  	CftC:=0.0
	else
  	CftC:=TobPiece.Getvalue('GP_TOTALHTDEV')/TobPiece.Getvalue('GP_MONTANTPA');
  if CftC < CftP then
  Begin
  	if AfficheMessage then
    Begin
    	Mess:='Vous êtes en dessous du coefficient de marge autorisé.';
			PgiBox(TraduireMemoire(Mess));
    End;
  	if Not V_PGI.Superviseur then result:=False;
  End;
end;

function ControleVisaBTP(TobPiece : TOB) : Boolean;
Var MontantVisa : Double;
    MontantPiece: Double;
    OkVisa      : Boolean;
    NaturePiece : string;
    EtatVisa    : string;
begin

  Result := True;

  NaturePiece := TobPiece.GetValue('GP_NATUREPIECEG');
  MontantPiece:= TobPiece.GetValue('GP_TOTALHT');
  Etatvisa    := TobPiece.GetValue('GP_ETATVISA');
  //
  if EtatVisa = 'NON' then Exit;
  //
  MontantVisa :=  GetInfoParPiece(NaturePiece, 'GPP_MONTANTVISA');
  OkVisa      := (GetInfoParPiece(NaturePiece, 'GPP_VISA')='X');

  if OkVisa then
  begin
    if (GetInfoParPiece(NaturePiece,'GPP_VISAIMP') = 'X') AND (EtatVisa <> 'VIS') then Result := False
  end;

end;

procedure SetTiersTTC (FF : Tform);
var XX : TFFacture;
    //newNature : string;
begin
  XX := TFFacture(FF);
  if XX.SaisieTypeAvanc = True then
  begin
    XX.gs.ListeParam := GetParamsoc('SO_LISTEAVANC');
  end else
  begin
    {$IFDEF BTP}
      XX.GS.ListeParam := GetInfoParPiece(XX.NewNature, 'GPP_LISTESAISIE');
    {$ELSE}
      XX.GS.ListeParam := GetInfoParPiece(XX.NewNature, 'GPP_LISTEAFFAIRE');
    {$ENDIF}
  end;
  XX.LesColonnes := XX.GS.titres[0];
end;


procedure PositionneModifFacturation(TOBL : TOB; DEV : Rdevise);
var QteDejaFact,MtDejaFac,PourcentAvanc : double;
begin
  MtDejaFac := Arrondi(TOBL.GEtValue('GL_QTEPREVAVANC')* TOBL.GEtValue('GL_PUHTDEV') / TOBL.GEtValue('GL_PRIXPOURQTE'),DEV.decimale);

  PourcentAvanc := Arrondi(MtDejaFac / TOBL.getValue('GL_MONTANTHTDEV') * 100,2);
  QteDejaFact := Arrondi(TOBL.getValue('GL_QTEFACT')* PourcentAvanc/100,V_PGI.okdecQ);
  TOBL.putValue('BLF_MTMARCHE',TOBL.getValue('GL_MONTANTHTDEV'));
  TOBL.putValue('BLF_MTDEJAFACT',MtDejaFac);
  TOBL.putValue('BLF_MTCUMULEFACT',MtDejaFac);
  TOBL.putValue('BLF_MTSITUATION',TOBL.GetValue('GL_MONTANTHTDEV'));
  TOBL.PutValue('BLF_MTPRODUCTION',TOBL.GetValue('BLF_MTSITUATION'));
  TOBL.putValue('BLF_QTEMARCHE',TOBL.getValue('GL_QTEFACT'));
  TOBL.putValue('BLF_QTEDEJAFACT',QteDejaFact);
  TOBL.putValue('BLF_QTECUMULEFACT',QteDejaFact);
  TOBL.putValue('BLF_QTESITUATION',0);
  TOBL.putValue('BLF_POURCENTAVANC',PourcentAvanc);
end;

procedure PositionneAncienneFacturation(TOBL : TOB; DEV : Rdevise; Typefacturation : string);
var QteMarche,MtMarche,PU,QteDejaFact,MtDejaFact,QteCumuleFact,MtCumuleFact,PourcentAvanc,PrixPour : double;
		QteSit,MtSit : double;
    QQ : TQuery;
    Sql : string;
    cledoc : R_CLEDOC;
begin
  if Typefacturation = 'DIR' then
  begin
    if TOBL.getValue('GL_PIECEPRECEDENTE') <> '' then
    begin
      // dans ce cas seul la quantité de la ligne de devis est importante
      // (ancienne gestion de facturation sur facture directe )
        //
      QteMarche := TOBL.GetValue('GL_QTEPREVAVANC');
      QteSit := TOBL.GetValue('GL_QTEFACT');
    	Pu := TOBL.GEtValue('GL_PUHTDEV');
      PrixPour := TOBL.GetValue('GL_PRIXPOURQTE'); if PrixPour = 0 then PrixPour := 1;
      MtMarche := Arrondi(QteMarche*Pu/PrixPour ,DEV.decimale);
      MtSit := Arrondi(QteSit*Pu/PrixPour ,DEV.decimale);
      TOBL.PutValue('BLF_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'));
      TOBL.PutValue('BLF_MTMARCHE',MtMarche );
      TOBL.PutValue('BLF_QTEMARCHE',QteMarche);
      TOBL.PutValue('BLF_MTSITUATION',MtSit );
      TOBL.PutValue('BLF_QTESITUATION',QteSit );
      if MtMarche <> 0 then TOBL.PutValue('BLF_POURCENTAVANC',Arrondi(MtSIT/MtMarche*100,2) )
      								 else TOBL.PutValue('BLF_POURCENTAVANC',0 )
    end;
  end else
  begin
    PrixPour := TOBL.GEtValue('GL_PRIXPOURQTE');if PrixPour = 0 then PrixPour := 1;
    Pu := TOBL.GEtValue('GL_PUHTDEV');
    //
    QteMarche := TOBL.GetValue('GL_QTEPREVAVANC');
    MtMarche := Arrondi(QteMarche*Pu/PrixPour ,DEV.decimale);
    //
    QteCumuleFact := TOBL.GetValue('GL_QTESIT');
    MtCumuleFact := Arrondi(QteCumuleFact*Pu/PrixPour,DEV.decimale);
    //
    QteSit := TOBL.GetValue('GL_QTEFACT');
    MTSit := TOBL.GEtValue('GL_MONTANTHTDEV');
    //
    QteDejaFact := Abs(TOBL.GetValue('GL_QTESIT')) - Abs(TOBL.GetValue('GL_QTEFACT'));
    MtDejaFact := Arrondi(QteDejaFact*Pu/PrixPour,DEV.decimale);
    //
    if MtMarche <> 0 then PourcentAvanc := Arrondi(MtCumuleFact/MtMarche* 100,2) else PourcentAvanc := 0; 
    //
    TOBL.putValue('BLF_MTMARCHE',MtMarche);
    TOBL.putValue('BLF_QTEMARCHE',QteMarche);
    //
    TOBL.putValue('BLF_MTDEJAFACT',MtDejaFact);
    TOBL.putValue('BLF_QTEDEJAFACT',QteDejaFact);
    //
    TOBL.putValue('BLF_MTCUMULEFACT',MtCumuleFact);
    TOBL.putValue('BLF_QTECUMULEFACT',QteCumuleFact);
    //
    TOBL.putValue('BLF_MTSITUATION',MtSit);
    TOBL.putValue('BLF_QTESITUATION',QteSit);
    //
    TOBL.putValue('BLF_POURCENTAVANC',PourcentAvanc);
  end;
end;

procedure InitMontantfacturation(TOBL : TOB);
begin
  TOBL.putValue('BLF_MTMARCHE',TOBL.getValue('GL_MONTANTHTDEV'));
  TOBL.putValue('BLF_MTDEJAFACT',0);
  TOBL.putValue('BLF_MTCUMULEFACT',0);
  TOBL.putValue('BLF_MTSITUATION',0);
  TOBL.putValue('BLF_QTEMARCHE',TOBL.getValue('GL_QTEFACT'));
  TOBL.putValue('BLF_QTEDEJAFACT',0);
  TOBL.putValue('BLF_QTECUMULEFACT',0);
  TOBL.putValue('BLF_QTESITUATION',0);
  TOBL.putValue('BLF_POURCENTAVANC',0);
end;

procedure PositionneAncienneFacDetOuv(TOBL,TOBOuvrages: TOB; DEV : RDevise ;Typefacturation : string);
var IndiceNomen,Indice : Integer;
		TOBOUV,TOBOO : TOB;
    Sql : string;
    QQ: TQuery;
    cledoc : R_CLEDOC;
    Qte,QteDev,QteOri,QteDuDetail,MtMarche,QteSit,MtSit,QteMarche : Double;
begin
	IndiceNomen := TOBL.GetValue('GL_INDICENOMEN'); if IndiceNomen = 0 then Exit;
  if TOBOuvrages.detail.Count = 0 then Exit;
  TOBOUV := TOBOuvrages.detail[IndiceNomen-1];
  QteOri := TOBL.GetValue ('GL_QTEPREVAVANC');   // qte du marche dans le devis
  QteDev := TOBL.GetValue('GL_QTEFACT');

  for Indice := 0 to TOBOUV.Detail.count -1 do
  begin
    TOBOO := TOBOUV.detail[Indice];
    QteDuDetail := TOBOO.getValue('BLO_QTEDUDETAIL'); if QteDuDetail = 0 then QteDuDetail := 1;
    Qte := TOBOO.GetValue('BLO_QTEFACT');
    //
    QteSit := Arrondi(Qte*QteDev / QteDuDetail,V_PGI.okdecQ);
    MtSit := Arrondi(QteSit*TOBOO.GetValue('BLO_PUHTDEV'),DEV.decimale);

    if TOBOO.GetValue('BLO_PIECEPRECEDENTE')<> '' then
    begin
      // nouvelle gestion avec gestion des sous detail pour facturation directe
      DecodeRefPieceOUv(TOBOO.GetValue('BLO_PIECEPRECEDENTE'),cledoc);
      if cledoc.UniqueBlo <> 0 then
      begin
        Sql := 'SELECT BLF_MTMARCHE,BLF_QTEMARCHE FROM LIGNEFAC WHERE '+ WherePiece(cledoc,ttdLignefac,false)+
               ' AND BLF_UNIQUEBLO='+inttoStr(Cledoc.UniqueBlo);
        QQ := OpenSQL(SQl,True,-1,'',true);
        if not QQ.eof then
        begin
          TOBOO.PutValue('BLF_NATUREPIECEG',TOBOO.GetValue('BLO_NATUREPIECEG'));
          TOBOO.PutValue('BLF_MTMARCHE',QQ.findField('BLF_MTMARCHE').AsFloat );
          TOBOO.PutValue('BLF_QTEMARCHE',QQ.findField('BLF_QTEMARCHE').AsFloat);
          TOBOO.PutValue('BLF_MTSITUATION',MtSit );
          TOBOO.PutValue('BLF_QTESITUATION',QteSit);
          if MtMarche > 0 then
            TOBOO.PutValue('BLF_POURCENTAVANC',ARRONDI(MtSit/MtMarche*100,DEV.Decimale))
          else
            TOBOO.PutValue('BLF_POURCENTAVANC',0);
        end;
        Ferme(QQ);
      end;
    end else
    begin
      if TOBL.getValue('GL_PIECEPRECEDENTE') <> '' then
      begin
        // dans ce cas seul la quantité de la ligne de devis est importante
        // (ancienne gestion de facturation sur facture directe )
        //
        QteMarche := Arrondi(Qte*QteOri / QteDuDetail,V_PGI.okdecQ);
        MtMarche := Arrondi(QteMarche*TOBOO.GetValue('BLO_PUHTDEV'),DEV.decimale);
        //
        TOBOO.PutValue('BLF_NATUREPIECEG',TOBOO.GetValue('BLO_NATUREPIECEG'));
        TOBOO.PutValue('BLF_MTMARCHE',MtMarche);
        TOBOO.PutValue('BLF_QTEMARCHE',QteMarche);
        TOBOO.PutValue('BLF_MTSITUATION',MtSit );
        TOBOO.PutValue('BLF_QTESITUATION',QteSit);
        if MtMarche > 0 then
          TOBOO.PutValue('BLF_POURCENTAVANC',ARRONDI(MtSit/MtMarche*100,DEV.Decimale))
        else
          TOBOO.PutValue('BLF_POURCENTAVANC',0);
      end;
    end;
  end;
end;


end.


