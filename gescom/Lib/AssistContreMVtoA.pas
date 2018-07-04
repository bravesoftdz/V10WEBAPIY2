unit AssistContreMVtoA;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
{$IFDEF EAGLCLIENT}
      maineagl,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
     UplannifChUtil,
     UTofListeInv,
     FactCommBtp,
     EntGc,
     UtilValidReapDecis,
{$ENDIF}
     Mask, HPanel,UTOB,HEnt1, FactTOB, FactPiece,
     ParamSoc, FactContreM, TntStdCtrls, TntComCtrls, TntExtCtrls ;

type
  TFCmdeContreM = class(TFAssist)

    TabSheet1: TTabSheet;
    TINTRO: THLabel;
    GBParam: TGroupBox;
    NbJourSecu: THNumEdit;
    HLabel1: THLabel;
    TabSheet2: TTabSheet;
    ListRecap: TListBox;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TRecap: THLabel;
    HLabel2: THLabel;
    PTITRE: THPanel;
    CBRECUPBLOCNOTE: TCheckBox;
    CBECLATCDEFOU: TCheckBox;

    procedure FormShow(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure createTOB;
    procedure DestroyTOB;
    Procedure ChargeTobLigSel;
    procedure bAnnulerClick(Sender: TObject);
  private
    { Déclarations privées }
    mrRes : boolean;
    TOBLigneSel: TOB ;
{$IFDEF BTP}
    procedure AjouteLigneEclate(TOBL : TOB);
    procedure ChargelesLignes(TOBLignes, TOBLIen, TheCurrent: TOB);
    procedure NettoyageTables(TOBLigne: TOB);

//    Deplace dans UtilValidReapDecis
//    procedure TrielaTOB(TOBLigneSel: TOB ; ParQuoi : string);
//    procedure ConstitueLaTOBFinale(TOBLigneSel: TOB; ChampsTri,RuptureDoc,ruptureArt,Mode : string);
//    procedure GenereLATobSortie(TOBIntermediaire, TOBLigneSel: TOB; Mode : string);
//    function ConstitueChampsTriRupt(var Sort, RuptDoc, RuptArt: string;EclatCdeFou: boolean): string;

{$ENDIF}
    function MiseAjourLigne (TOBL,TOBLS,TOBRef : TOB) : boolean;
  public
    { Déclarations publiques }
    TOBLigne,TOBCataLogu: TOB ;
    QteATraiterName, Mode : String;       // Nom du champ sup. quantité à commander

  end;

var
  FCmdeContreM: TFCmdeContreM;

implementation

{$R *.DFM}
uses StockUtil,FactGrp,FactUtil,FactComm,UtilBTPgestChantier, BTPUtil;
{===============================================================================}
{---------------------------- Gestion de la fiche ------------------------------}
{===============================================================================}
procedure TFCmdeContreM.FormCreate(Sender: TObject);
begin
  inherited;
  CreateTOB;
end;

procedure TFCmdeContreM.FormShow(Sender: TObject);
begin
  inherited;
{$IFDEF BTP}
CBRECUPBLOCNOTE.Visible := false;
CBECLATCDEFOU.Visible := true;
CBECLATCDEFOU.Checked := GetparamSoc('SO_BTECLATCDEFOU');
{$ENDIF}
bAnnuler.Visible := True;
bSuivant.Enabled := True;
bFin.Visible := True;
bFin.Enabled := False;
end;


procedure TFCmdeContreM.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
if mrRes then
   ModalResult:=mrOk
   else
   ModalResult:=mrCancel;

  DestroyTOB;
  inherited;
end;

procedure TFCmdeContreM.bSuivantClick(Sender: TObject);
begin
  inherited;
if (bSuivant.Enabled) then
   bFin.Enabled := False
   else
   begin
      bFin.Enabled := True;
      ListRecap.Clear;
      ListRecap.Items.Add('');
      ListRecap.Items.Add('Délai de sécurité : ' + NbJourSecu.Text + ' jour(s)');
   End;
end;

procedure TFCmdeContreM.bPrecedentClick(Sender: TObject);
begin
  inherited;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end;

{$IFDEF BTP}
(*
function TFCmdeContreM.ConstitueChampsTriRupt(var Sort: string; var RuptDoc : string; Var RuptArt : string; EclatCdeFou : boolean): string;  //a la place de tri j'ai faillit ecrire Tage
begin
//	IF CBECLATCDEFOU.checked then
	IF EclatCdeFou then
  begin
    Sort := 'GL_TIERS;GL_DEPOT;GL_AFFAIRE;GL_IDENTIFIANTWOL;GL_ARTICLE;';
    RuptDoc := 'GL_TIERS;GL_DEPOT;GL_AFFAIRE;GL_IDENTIFIANTWOL;';
    RuptArt := Sort;
  end else
  begin
    Sort := 'GL_TIERS;GL_DEPOT;GL_ARTICLE;GL_IDENTIFIANTWOL;GL_AFFAIRE';
    RuptDoc := 'GL_TIERS;';
    RuptArt := 'GL_TIERS;GL_DEPOT;GL_ARTICLE;GL_IDENTIFIANTWOL;';
  end;
end;

procedure TFCmdeContreM.TrielaTOB(TOBLigneSel: TOB ; ParQuoi : string);
begin
	TOBLigneSel.Detail.Sort(ParQuoi);
end;
procedure TFCmdeContreM.ConstitueLaTOBFinale (TOBLigneSel: TOB; ChampsTri,RuptureDoc,ruptureArt,Mode : string);

	function GetValeurElt(TOBL : TOB;lesChamps : string) : string;
  var UnChamps,UneValeur : string;
  begin
  	result := '';
    repeat
    	unChamps := readTokenSt (LesChamps);
      if TOBL.fieldExists (unChamps) then
      begin
      	result := result + VarasType(TOBL.getValue (UnChamps),varString)+';';
      end;
    until unchamps = '';
  end;

var TOBIntermediaire : TOB;
		TOBL,TOBLD,TOBDest,TOBDArt,TOBLART,TOBLReg,TOBDoc : TOB;
    Indice,Detail,NumOrdre,IArticle,Document: integer;
    TraitementDoc,TraitementArt,ComparaisonSort : string;
    ThisDoc,ThisArt,CurrentSort : string;
begin
	TOBIntermediaire := TOB.Create ('THE TOB INTERMEDIAIRE',nil,-1);
  TraitementDoc := '';
  TraitementArt := '';
  ComparaisonSort := '';
  Document := 1;
  TRY
    //
    NumOrdre := 1;
    for Indice := 0 to TOBLigneSel.detail.count -1 do
    begin
    	TOBL := TOBLigneSel.detail[Indice];
    	// Traitement de la rupture
      ThisDoc := GetValeurElt(TOBL,RuptureDoc);
      ThisArt := GetValeurElt(TOBL,RuptureArt);
      CurrentSort := GetValeurElt(TOBL,ChampsTri);
      if (TraitementDoc <> ThisDoc) then
      begin
      	// Emulation de la Creation du document
        TOBDest := TOB.Create ('UN DOCUMENT SORTIE',TOBIntermediaire,-1);
        TOBDest.AddChampSupValeur ('DOCUMENT',ThisDoc,false);
        TraitementDoc := ThisDoc;
        TraitementArt := '';
      end;
      if TraitementArt <> ThisArt then
      begin
        TOBDART := TOB.Create ('UN ARTICLE',TOBDest,-1);
        TOBDART.AddChampSupValeur ('ARTICLE',ThisArt,false);
        TraitementArt := ThisArt;
				TOBDART.AddChampSupValeur ('NUMRATACHEMENT',NumOrdre);
        inc(NumOrdre);
      end;
      TOBLART := TOB.Create ('LIGNE',TOBDART,-1);
      TOBLArt.Dupliquer (TOBL,true,true);
    end;
    //
    GenereLATobSortie_UtilVRD (TOBIntermediaire,TOBLigneSel,Mode);
  FINALLY
  	TOBInterMediaire.free;
  END;
end;

procedure TFCmdeContreM.GenereLATobSortie (TOBIntermediaire,TOBLigneSel : TOB; Mode : string);

  function IsMultiAffaire (LATOB : TOB) : boolean;
  var Indice : integer;
  		LaRef : string;
      ThisTOB : TOB;
  begin
  	result := false;
    for Indice := 0 to LaTOB.detail.count -1 do
    begin
    	ThisTOB := LaTOB.detail[Indice];
      if Indice = 0 then
      begin
      	LaRef := ThisTOB.GetValue('GL_AFFAIRE');
      end;
      if LaRef <> ThisTOB.GetValue('GL_AFFAIRE') then
      begin
      	result := true;
        break;
      end;
    end;
  end;

var NbrDoc,NbrArt,DetARt : integer;
		TOBDoc,TOBDDOC,TOBLD,TOBART,TOBLReg,TOBLART : TOB;
    Cledoc : R_cledoc;
begin
  TOBLigneSel.ClearDetail;
  // On réécrit dans la toblignesel avec les regroupements qui vont bien
  for NbrDoc := 0 to TOBIntermediaire.detail.count -1 do
  begin
    TOBDOC := TOBIntermediaire.detail[NbrDoc]; // Un document a traiter
    for NbrArt := 0 To TOBDOC.Detail.count -1 do
    begin
      TOBART := TOBDOC.Detail[NbrArt]; // Ligne de regroupement tOB intermediaire
      if TOBART.Detail.Count > 1 then
      begin
        TOBLReg := TOB.Create ('LIGNE',TOBLigneSel,-1); // Ca c'est la ligne de regroupement sur document destination
        for DetARt := 0 to TOBART.detail.count -1 do
        begin
          TOBLD := TOBART.detail[DetARt];
          if Detart = 0 then
          begin
            TOBLReg.Dupliquer (TOBLD,true,true);
            TOBLReg.AddChampSupValeur ('NUMRATACHEMENT',TOBART.GetValue('NUMRATACHEMENT'));
            TOBLReg.AddChampSupValeur ('DOCUMENT',TOBDOC.GetValue('DOCUMENT'));
            if (Mode = 'REAP') and (IsMultiAffaire(TOBLD)) then
            begin
              TOBLReg.putValue('GL_AFFAIRE','');
              TOBLReg.putValue('GL_AFFAIRE1','');
              TOBLReg.putValue('GL_AFFAIRE2','');
              TOBLReg.putValue('GL_AFFAIRE3','');
              TOBLReg.putValue('GL_AVENANT','');
            end;
            TOBLReg.putValue('GL_PIECEPRECEDENTE','');
            DecodeRefPiece (TOBLReg.GetValue('GL_PIECEORIGINE'),cledoc);
            cledoc.NumLigne := 0;
            cledoc.NumOrdre := 0;
            TOBLReg.putValue('GL_PIECEORIGINE',CleDocToString (cledoc));
//            TOBLReg.putValue('GL_PIECEORIGINE','');
            if TOBLReg.FieldExists ('PIECEORIGINE') then  TOBLReg.DelChampSup ('PIECEORIGINE',false);
            TOBLReg.putValue('GL_TYPELIGNE','CEN');
            TOBLReg.putValue('GL_LIGNELIEE',0);
            TOBLReg.putValue('GL_QTEFACT',0);
            TOBLReg.putValue('GL_TENUESTOCK','-');
          	TOBLReg.PutValue ('NUMRATACHEMENT',TOBART.GetValue('NUMRATACHEMENT'));
            TOBLReg.putValue('GL_LIGNELIEE',TOBLReg.GetValue('NUMRATACHEMENT'));
          end;

          TOBLArt := TOB.Create ('LIGNE',TOBLIgneSel,-1); // Et Là la ligne detail
          TOBLArt.dupliquer (TOBLD,true,true);
          TOBLART.AddChampSupValeur ('RATACHE',TOBArt.GEtValue('NUMRATACHEMENT'));
          TOBLART.AddChampSupValeur ('DOCUMENT',TOBDOC.GetValue('DOCUMENT'));
          TOBLART.putValue('GL_LIGNELIEE',TOBLART.GetValue('RATACHE'));
          // Cumul
          TOBLReg.putValue ('GL_QTEFACT',TOBLReg.GetValue ('GL_QTEFACT')+TOBLART.getValue('GL_QTEFACT'));
          TOBLReg.putValue ('GL_QTESTOCK',TOBLReg.GetValue ('GL_QTEFACT'));
          TOBLReg.putValue ('GL_QTERESTE',TOBLReg.GetValue ('GL_QTEFACT'));
        end;
      end else
      begin
      	TOBLD := TOBART.detail[0];
        TOBLArt := TOB.Create ('LIGNE',TOBLIgneSel,-1); // Et La la ligne detail
        TOBLArt.dupliquer (TOBLD,true,true);
        TOBLART.AddChampSupValeur ('DOCUMENT',TOBDOC.GetValue('DOCUMENT'));
      end;
    end;
  end;
end;
*)

{$ENDIF}

procedure TFCmdeContreM.bFinClick(Sender: TObject);
var 
{$IFDEF BTP}
		Tri,RuptureDoc,RuptureArt : string;
{$ENDIF}
begin
  inherited;
  mrRes := false ;
{$IFDEF BTP}
  if (CBECLATCDEFOU.Checked) and (Mode='REAP') then Mode := 'REAPECLAT';
{$ENDIF}
  ChargeTobLigSel;
// AJOUT LS Pour Nouvelle gestion
{$IFDEF BTP}
  ConstitueChampsTriRupt_UtilVRD (Tri,RuptureDoc,RuptureArt,CBECLATCDEFOU.checked);
	TrielaTOB_UtilVRD(TOBLigneSel,Tri);
  ConstitueLaTOBFinale_UtilVRD (TOBLigneSel,Tri,RuptureDoc,RuptureArt,Mode,StrToInt(FloatToStr(NbJourSecu.value)));
{$ENDIF}
// --
  if TOBLigneSel.detail.count>0 then
  begin
       //Genere les pieces
       mrRes:= CreerPiecesFromLignes(TobLigneSel, Mode,iDate1900);
{$IFDEF BTP}
			 if mrRes Then NettoyageTables(TOBLigne);
{$ENDIF}
  end;
Close ;
end;

{===============================================================================}
{------------------------------ Gestion des TOBs -------------------------------}
{===============================================================================}
procedure TFCmdeContreM.CreateTOB ;
begin
TOBLigneSel:=TOB.Create('LIGNE',Nil,-1) ;
TOBCataLogu := TOB.Create ('LES CATALOG', nil,-1);
end;

{$IFDEF BTP}
Procedure TFCmdeContreM.ChargelesLignes (TOBLignes,TOBLIen,TheCurrent : TOB);
var TOBL,TOBLI,TOBLS : TOB;
    Indice : integer;
    cledoc : CleligneDevCha;
    Requete : string;
    QQ : TQuery;
    FUS,FUA,FUV,QTeCumul,TheQteInterm,QteBesoin,CoefUsUv : double;
    MtBesoin      : Double;
    MtCumul       : Double;
    TheMtInterm   : Double;
    Ok_ReliquatMt : Boolean;
begin
  (* TOBL = Ligne de commande EN UV*)
  (* TheCurrent = Ligne de reapro EN US *)
  QteBesoin := TheCurrent.getValue('GL_QTEFACT');   // EN US Donc
  // --- GUINIER ---
  Ok_ReliquatMt := CtrlOkReliquat(TheCurrent, 'GL');
  MtBesoin := TheCurrent.GetValue('GL_MONTANTHTDEV');

  QteCumul  := 0;
  MtCumul   := 0;

  for Indice := 0 to TOBLien.detail.count -1 do
  begin
    if Ok_ReliquatMt then
      if MtCumul  >= MtBesoin then Break // BRL090608 : la quentité cumulée couvre déjà le besoin  : on sort (cf FQ11789 pour CHS)
    else
      if QteCumul >= QteBesoin then Break; // BRL090608 : la quentité cumulée couvre déjà le besoin  : on sort (cf FQ11789 pour CHS)

    TOBLI := TOBLien.detail[Indice];

    cledoc := DecodeLienDevCHA (TOBLI.GetValue('BDA_REFD'));
(*    requete := MakeSelectLigneBtp (True,True); *)

		Requete := 'SELECT LIGNE.*,LIGNEPHASES.BLP_PHASETRA '+
               'FROM LIGNE '+
               'LEFT JOIN LIGNEPHASES ' +
               'ON (BLP_NATUREPIECEG = GL_NATUREPIECEG and AND BLP_SOUCHE=GL_SOUCHE AND BLP_NUMERO = GL_NUMERO ' +
               'and BLP_INDICEG = GL_INDICEG and BLP_NUMORDRE = GL_NUMORDRE)';
    requete := requete + ' WHERE GL_NATUREPIECEG="'+cledoc.NaturePiece+'" AND ';
    Requete := requete + 'GL_SOUCHE="'+cledoc.Souche+'" AND ';
    Requete := requete + 'GL_NUMERO='+IntToStr(cledoc.NumeroPiece) +' AND ';
    Requete := requete + 'GL_INDICEG='+IntToStr(cledoc.Indice) +' AND ';
    Requete := requete + 'GL_NUMLIGNE='+IntToStr(TOBLI.GetValue('BDA_NUMLD'));
    QQ := OpenSql (Requete,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBL := TOB.Create ('LIGNE',TOBLignes,-1);
      TRY
        TOBL.SelectDB ('',QQ);
        TOBLS:=NewTOBLigne(TOBLigneSel,-1);
        TOBLS.Dupliquer (TheCurrent,false,true);
        // transformation en Unite d'achat (Les elements sont en unites de ventes)
        FUV := RatioMesure('PIE', TobLS.GetValue('GL_QUALIFQTEVTE'));
        FUS := RatioMesure('PIE', TobLS.GetValue('GL_QUALIFQTESTO'));
        FUA := RatioMesure('PIE', TobLS.GetValue('GL_QUALIFQTEACH'));
        if FUV = 0 then FUV := 1;
        CoefUsUv := TOBL.GetDouble('GL_COEFCONVQTEVTE');
        if CoefUsUv = 0 then CoefusUv := (FUS/FUV);
// RESOLUTION DU CAS OU LA QUANTITE DEVIENT SUPERIEURE A CELLE DU BESOIN REEL
// Ca c'etait avantttt				TheQteInterm := (TobL.GetValue('GL_QTEFACT') * FUV)/FUS; // en US

//				if TOBLI.GetValue('BDA_QUANTITE') = 0 then TobLi.PutValue('BDA_QUANTITE',(TobL.GetValue('GL_QTEFACT') * FUV)/FUS);
				if TOBLI.GetValue('BDA_QUANTITE') = 0 then TobLi.PutValue('BDA_QUANTITE',(TobL.GetValue('GL_QTEFACT') / CoefusUV));
//			TheQteInterm := (TobLi.GetValue('BDA_QUANTITE') * FUV)/FUS; // en US --> passage en UV

        // --- GUINIER ---
        if Ok_ReliquatMt then
        begin
          TheMtInterm := TobLi.GetValue('BDA_QUANTITE') * TOBLS.GetValue('GL_PUHTDEV') ; // en US --> passage en UV
          if TheMtInterm  > MtBesoin - MtCumul then
            TobLS.PutValue('GL_MTRESTE', MtBesoin - MtCumul)
          else
            TobLS.PutValue('GL_MTRESTE', TheMtInterm);
        end
        else
        begin
          TheQteInterm := TobLi.GetValue('BDA_QUANTITE') ; // en US --> passage en UV
          if TheQteInterm  > QteBesoin - QteCumul then
          begin
            TobLS.PutValue('GL_QTEFACT', QteBesoin - QteCumul);
            TOBLS.PutValue('GL_QTESTOCK',TobLS.GetValue('GL_QTEFACT'));
            TOBLS.PutValue('GL_QTERESTE',TobLS.GetValue('GL_QTEFACT'));
          end
          else
          begin
  //          TobLS.PutValue('GL_QTEFACT', (TobL.GetValue('GL_QTESTOCK') * FUV)/FUS);
            TobLS.PutValue('GL_QTEFACT', TheQteInterm);
            TOBLS.PutValue('GL_QTESTOCK',TobLS.GetValue('GL_QTEFACT'));
            TOBLS.PutValue('GL_QTERESTE',TobLS.GetValue('GL_QTEFACT'));
          end;
        end;

        TOBLS.PutValue('GL_DEPOT',TobL.GetValue('GL_DEPOT'));

        // Recup de l'affaire initiale
        TobLS.PutValue('GL_AFFAIRE', TobL.GetValue('GL_AFFAIRE'));
        TobLS.PutValue('GL_AFFAIRE1', TobL.GetValue('GL_AFFAIRE1'));
        TobLS.PutValue('GL_AFFAIRE2', TobL.GetValue('GL_AFFAIRE2'));
        TobLS.PutValue('GL_AFFAIRE3', TobL.GetValue('GL_AFFAIRE3'));

        // traçabilité pièce
        TOBLS.PutValue('GL_PIECEPRECEDENTE','') ;
        TOBLS.PutValue('GL_PIECEORIGINE',EncodeRefPiece(TOBL)) ;
        TOBLS.addChampSupValeur('ORIGINE',EncodeRefPiece(TOBL,0,false));
        // Recup destination livraison fournisseur
        TOBLS.putValue('GL_IDENTIFIANTWOL',TOBL.GetValue('GL_IDENTIFIANTWOL'));
        //
        (* CONSO *)
        if VarIsNull(TOBL.getValue('BLP_PHASETRA')) or (VarAsType(TOBL.getValue('BLP_PHASETRA'), varString) = #0) then
          TOBL.PutValue('BLP_PHASETRA','');
        TOBLS.AddChampSupValeur ('BLP_PHASETRA',TOBL.getValue('BLP_PHASETRA'));
        (* ----- *)
        TOBLS.PutValue ('GL_DATELIVRAISON',QQ.findfield('GL_DATELIVRAISON').asdatetime);
//        QteCumul := QteCumul + (TOBL.GetValue('GL_QTEFACT') * FUV / FUS ); // Cumul en US
        QteCumul := QteCumul + TheQteInterm; // Cumul en US
//        MiseAjourLigne (TOBL,TOBLS,TheCurrent);
        MiseAjourLigne (TOBL,TOBLS,TOBLS);
        // CORRECTION LS POUR REAPPRO
        TOBLS.PutValue('GL_TENUESTOCK',TheCurrent.getValue('GL_TENUESTOCK'));
        // --
      FINALLY
        TOBL.free;
      END;
    end;
    ferme (QQ);
  end;
  // le reste non relié a des documents de ventes proviennent forcemment de demande pour le stock
  if QteCumul < (TheCurrent.getValue('GL_QTEFACT')) then
  begin
    TOBLS:=NewTOBLigne(TOBLigneSel,-1);
    TOBLS.Dupliquer (TheCurrent,false,true);
    FUA := RatioMesure('PIE', TheCurrent.GetValue('GL_QUALIFQTEACH'));
    FUS := RatioMesure('PIE', TheCurrent.GetValue('GL_QUALIFQTESTO'));
    TOBLS.PutValue('GL_QTEFACT',(TheCurrent.getValue('GL_QTEFACT') - QteCumul) * (FUS/FUA));
    TOBLS.PutValue('GL_QTESTOCK',TOBLS.GetValue('GL_QTEFACT'));
    TOBLS.PutValue('GL_QTERESTE',TOBLS.GetValue('GL_QTEFACT'));
    TOBLS.PutValue('GL_PIECEPRECEDENTE','') ;
    TOBLS.PutValue('GL_PIECEORIGINE',EncodeRefPiece(TheCurrent)) ; // piece origine = reappro
    (* CONSO *)
    TOBLS.AddChampSupValeur ('BLP_PHASETRA','');
    (* ----- *)
    MiseAjourLigne (TOBLS,TOBLS,TheCurrent);

  end;
  // --- GUINIER ---
  if CtrlOkReliquat(TOBLS, 'GL') then TOBLS.PutValue('GL_MTRESTE',TOBLS.GetValue('GL_MONTANTHTDEV'));

end;

Procedure TFCmdeContreM.AjouteLigneEclate (TOBL : TOB);
var QQ,QU : TQuery;
    TheCurrent,TOBLien,TOBLignes,TOBLS : TOB;
    Requete : string;
begin
  TOBLien := TOB.Create ('LES LIENS',nil,-1);
  TheCurrent := TOB.Create ('LIGNE',nil,-1);
  TOBLignes := TOB.Create ('LES LIGNES DE COMMANDES',nil,-1);
  QQ := OpenSQL('SELECT * FROM LIGNE WHERE ' + TOBL.cle1,True,-1,'',true) ;  // Les lignes de la reappro.
  TRY
    if not QQ.eof then
    begin
      TheCurrent.SelectDB ('',QQ,true);
      Requete := 'SELECT * FROM LIENDEVCHA WHERE BDA_REFC="'+EncodeLienDevCHA(TheCurrent)+'" AND ';
      Requete := Requete + 'BDA_NUMLC='+inttostr(TheCurrent.GetValue('GL_NUMLIGNE'));
      QU := OpenSql(requete,true,-1,'',true);
      if not QU.eof then
      begin
        TOBLien.LoadDetailDB ('LIENDEVCHA','','',QU,false);
        ChargelesLignes (TOBLignes,TOBLIen,TheCurrent);
      end else
      begin
        TOBLS:=NewTOBLigne(TOBLigneSel,-1);
        TOBLS.dupliquer (TheCurrent,false,true);
        TOBLS.PutValue('GL_PIECEPRECEDENTE','') ;
        TOBLS.PutValue('GL_PIECEORIGINE',EncodeRefPiece(TheCurrent)) ;
        (* CONSO *)
        TOBLS.AddChampSupValeur('BLP_PHASETRA','');
        (* ----- *)
        TOBLS.addChampSupValeur('ORIGINE',EncodeRefPiece(TheCurrent,0,false));
        MiseAjourLigne (TOBL,TOBLS,TheCurrent);
      end;
      ferme (QU);
    end;
  FINALLY
    ferme (QQ);
    TheCurrent.free;
    TOBLIen.free;
    TOBLIgnes.free;
  END;
end;
{$ENDIF}

function TFCmdeContreM.MiseAjourLigne (TOBL,TOBLS,TOBRef : TOB) : boolean;
var sTmp:string;
    Qte:double;
    DateL:TDateTime;
    DelaiFou:integer;
    TOBCata:TOB;
    QQ : TQuery;
    Select : string;
    TheStrDate : String;
begin
  result := true;
  //Mise à jour des tiers et fournissuer
  sTmp:=TOBLS.GetValue('GL_TIERS');
  TOBLS.PutValue('GL_TIERS',TOBRef.getValue('GL_FOURNISSEUR'));
  if TOBLS.GetValue('GL_TIERS')='' then begin result := false; end;
  TOBLS.PutValue('GL_TIERSFACTURE',TOBRef.getValue('GL_FOURNISSEUR'));
  TOBLS.PutValue('GL_FOURNISSEUR',sTMP);
  TOBLS.PutValue('GL_TIERSPAYEUR','');
  TOBLS.PutValue('GL_TIERSLIVRE','');

  if Trim(TOBLS.GetValue('GL_DEPOT')) = '' then
     TOBLS.PutValue('GL_DEPOT', GetParamSoc('SO_GCDEPOTDEFAUT'));

  if Not CBRECUPBLOCNOTE.Checked then
  begin
    TOBLS.PutValue('GL_BLOCNOTE','');
  end else
  begin
    TOBLS.PutValue('GL_BLOCNOTE',TOBRef.GetValue('GL_BLOCNOTE'));
  end;

  //Maj des qté
  //Déjà converti en unité de stock dans la fiche de sélection
  Qte:=TOBLS.getValue(QteATraiterName);
  TOBLS.putValue('GL_QTESTOCK',Qte);
  TOBLS.putValue('GL_QTEFACT', Qte);
  TOBLS.putValue('GL_QTERELIQUAT',0);
  TOBLS.putValue('GL_QTERESTE',Qte);  { NEWPIECE }
  // --- GUINIER ---
  if CtrlOkReliquat(TOBLS, 'GL') then
  begin
    TOBLS.putValue('GL_MTRELIQUAT',0);
    TOBLS.putValue('GL_MTRESTE', TOBLS.GetValue('GL_MONTANTHTDEV'));  { NEWPIECE }
  end;

   //Maj des prix
  if (Mode <> 'REAP') and (Mode <> 'REAPECLAT') then
  begin
    TOBLS.putValue('GL_PUHT',   TOBRef.getValue('GL_DPA'));
    TOBLS.putValue('GL_PUHTDEV',TOBRef.getValue('GL_DPA'));
  end else
  begin
    TOBLS.putValue('GL_QUALIFQTEVTE',TOBRef.getValue('GL_QUALIFQTEACH'));
  end;
  // si prix d'achat à 0, reprise du prix budget
  if TOBLS.getValue('GL_PUHTDEV') = 0 then
  begin
     TOBLS.PutValue('GL_PUHTDEV',TOBRef.getValue('GL_PUHTBASE'));
     TOBLS.PutValue('GL_PUHT',TOBRef.getValue('GL_PUHTBASE'));
  end;

  //Raz Remise
  TOBLS.putValue('GL_ESCOMPTE',0);
  TOBLS.putValue('GL_REMISELIGNE',0);

   //Délai
{$IFDEF BTP}
  TOBCata:=TOBCatalogu.FindFirst(['GCA_ARTICLE', 'GCA_TIERS'],[TOBLS.GetValue('GL_ARTICLE'),TOBLS.GetValue('GL_TIERS') ], False) ;
{$ELSE}
  TOBCata:=TOBCatalogu.FindFirst(['GCA_REFERENCE', 'GCA_TIERS'],[TOBLS.GetValue('GL_REFCATALOGUE'), GetCodeFourDCM(TOBL)], False) ;
{$ENDIF}
  if TOBCata = nil then
  begin
    Select := 'SELECT * FROM CATALOGU WHERE GCA_ARTICLE="'+TOBLS.GetValue('GL_ARTICLE')+'" AND '+
{$IFDEF BTP}
              'GCA_TIERS="'+TOBLS.GetValue('GL_TIERS')+'"';
{$ELSE}
              'GCA_TIERS="'+GetCodeFourDCM(TOBL)+'"';
{$ENDIF}
    QQ := OpenSql (Select,true,-1,'',true);
    if not QQ.eof then
    begin
      TOBCata := TOB.Create ('CATALOGU',TOBCatalogu,-1);
      TOBCata.SelectDB ('',QQ,true);
    end;
    Ferme (QQ);
  end;
  if TobCata<>nil then
  begin
   DelaiFou:=TOBCata.GetValue('GCA_DELAILIVRAISON') ;
{$IFDEF BTP}
   TOBLS.PutValue('GL_REFARTSAISIE',TOBCata.GetValue('GCA_REFERENCE')) ;
{$ENDIF}
{      if TOBCata.GetValue('GCA_PRIXBASE')>0 then   ///  BBI : Faudra qu'on m'explique....
         begin
         TOBLS.PutValue('GL_PUHT',TOBCata.GetValue('GCA_PRIXBASE'));
         TOBLS.PutValue('GL_PUHTDEV',TOBCata.GetValue('GCA_PRIXBASE'));
         end; }
  end else DelaiFou:=0;
  TheStrDate := DateToStr(TOBRef.GetValue('GL_DATELIVRAISON'));
  DateL:=TOBRef.GetValue('GL_DATELIVRAISON')-strtoint(NbJourSecu.Text)-DelaiFou;
//  if DateL < 2 then DateL:=2;
  if DateL < Date then TOBLS.putValue('GL_DATELIVRAISON',Date);
  TheStrDate := DateToStr(TOBRef.GetValue('GL_DATELIVRAISON'));

  //--- GUINIER ---
  if CtrlOkReliquat(TOBLS, 'GL') then
    TOBLS.putValue('GL_MTRESTE', Qte * TOBLS.GetValue('GL_PUHTDEV'));  { NEWPIECE }

end;

procedure TFCmdeContreM.DestroyTOB ;
begin
TOBLigneSel.Free ; TOBLigneSel := Nil;
TOBCataLogu.free;
end;

Procedure TFCmdeContreM.ChargeTobLigSel;
Var TOBL,TOBLS,TOBCata:TOB;
    i:integer;
    sTmp,RefPiece:string;
    Qte:double;
    DateL:TDateTime;
    DelaiFou:integer;
    Q:TQuery;
begin
  TOBLS := nil ;
  TOBLigneSel.ClearDetail ;
  if TOBLigne.detail.count<=0 then exit;
  // Génére la Tob des lignes sélectionnées
  for i:=0 to TOBLigne.Detail.Count - 1 do
     BEGIN
     TOBL:=TOBLigne.Detail[i];
     if TOBL.GetValue('SELECT') <> 'X' then continue;
  {$IFDEF BTP}
       // Ajouter eclatement par affaire et lieu de livraison
       AjouteLigneEclate (TOBL);
  {$ELSE}
       Q := OpenSQL('Select * from Ligne Where ' + TOBL.cle1,True,-1,'',true) ;
       if Not Q.EOF then
       BEGIN
          TOBLS:=NewTOBLigne(TOBLigneSel,-1);
          TOBLS.SelectDB('',Q);
          if not MiseAjourLigne (TOBL,TOBLS,TOBLS) then BEGIN TOBLS.Free; Continue; END;
       END;
       Ferme (Q);
{$ENDIF}
  (*   Q := OpenSQL('SELECT * FROM LIGNE WHERE GL_NATUREPIECEG="'+TOBL.GetValue('GL_NATUREPIECEG')+'"'+
       ' AND GL_SOUCHE="'+TOBL.GetValue('GL_SOUCHE')+ '" AND GL_NUMERO='+IntToStr(TOBL.GetValue('GL_NUMERO'))+
       ' AND GL_INDICEG='+IntToStr(TOBL.GetValue('GL_INDICEG'))+ ' AND GL_NUMLIGNE='+IntToStr(TOBL.GetValue('GL_NUMLIGNE')),True) ;
     if Not Q.EOF then
        BEGIN
        TOBLS:=NewTOBLigne(TOBLigneSel,-1);
        TOBLS.SelectDB('',Q);
        Ferme(Q) ;
        END else
        BEGIN
        Ferme(Q) ;
        Continue;
        END;

     //Mise à jour des tiers et fournissuer
     sTmp:=TOBLS.GetValue('GL_TIERS');
     TOBLS.PutValue('GL_TIERS',TOBLS.getValue('GL_FOURNISSEUR'));
     if TOBLS.GetValue('GL_TIERS')='' then begin TOBLS.Free; continue; end;
     TOBLS.PutValue('GL_TIERSFACTURE',TOBLS.getValue('GL_FOURNISSEUR'));
     TOBLS.PutValue('GL_FOURNISSEUR',sTMP);
     TOBLS.PutValue('GL_TIERSPAYEUR','');
     TOBLS.PutValue('GL_TIERSLIVRE','');
     if Not CBRECUPBLOCNOTE.Checked then TOBLS.PutValue('GL_BLOCNOTE','');

     if Trim(TOBLS.GetValue('GL_DEPOT')) = '' then
       TOBLS.PutValue('GL_DEPOT', GetParamSoc('SO_GCDEPOTDEFAUT'));

     //Maj des qté
     //Déjà converti en unité de stock dans la fiche de sélection
     Qte:=TOBL.getValue(QteATraiterName);
     TOBLS.putValue('GL_QTESTOCK',Qte);
     TOBLS.putValue('GL_QTEFACT',Qte);
     TOBLS.putValue('GL_QTERELIQUAT',0);
     TOBLS.putValue('GL_QTERESTE', Qte);  { NEWPIECE }
     //Maj des prix
     if Mode <> 'REAP' then
     begin
       TOBLS.putValue('GL_PUHT',TOBLS.getValue('GL_DPA'));
       TOBLS.putValue('GL_PUHTDEV',TOBLS.getValue('GL_DPA'));
     end;

     //Raz Remise
     TOBLS.putValue('GL_ESCOMPTE',0);
     TOBLS.putValue('GL_REMISELIGNE',0);

     //Délai
     TOBCata:=TOBCatalogu.FindFirst(['GCA_REFERENCE', 'GCA_TIERS'],[TOBLS.GetValue('GL_REFCATALOGUE'), GetCodeFourDCM(TOBLigne.Detail[i])], False) ;
     if TobCata<>nil then
        begin
        DelaiFou:=TOBCata.GetValue('GCA_DELAILIVRAISON') ;
  {      if TOBCata.GetValue('GCA_PRIXBASE')>0 then   ///  BBI : Faudra qu'on m'explique....
           begin
           TOBLS.PutValue('GL_PUHT',TOBCata.GetValue('GCA_PRIXBASE'));
           TOBLS.PutValue('GL_PUHTDEV',TOBCata.GetValue('GCA_PRIXBASE'));
           end; }
        end
        else
        DelaiFou:=0;
     DateL:=TOBLS.GetValue('GL_DATELIVRAISON')-strtoint(NbJourSecu.Text)-DelaiFou;
     if DateL < 2 then DateL:=2;
     TOBLS.putValue('GL_DATELIVRAISON',DateL);
     // traçabilité pièce
     RefPiece:=EncodeRefPiece(TOBL);
     TOBLS.PutValue('GL_PIECEPRECEDENTE','') ;
     TOBLS.PutValue('GL_PIECEORIGINE',RefPiece) ;

     END;
  // Fin
  *)
     END;
end;


procedure TFCmdeContreM.bAnnulerClick(Sender: TObject);
begin
  inherited;
  ModalResult:=mrCancel;
end;

{$IFDEF BTP}
procedure TFCmdeContreM.NettoyageTables(TOBLigne : TOB);
begin
  ExecuteSQL ('DELETE FROM LIENDEVCHA WHERE BDA_REFC LIKE "'+EncodeLienDevCHADel(TOBLigne)+'"');
end;
{$ENDIF}

end.
