unit UtilGenContreEtude;

interface

Uses Classes,
     Windows,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MainEagl,
{$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls,
     Hpanel,
     HEnt1,
     HMsgBox,
     Vierge,
     ExtCtrls,
		 AglInit,
     ComObj,
     UTOB,
     PiecesRecalculs;

procedure DefinitPieceCtrEtude ( TOBdevis,TOBCtrEtude,TOBOuvrage,TOBOuvrageResult,TOBInterDoc : TOB ; RootCtrEtude : TTreeNode);

implementation
uses factTOB,FactVariante,FactOuvrage,FactUtil,FactComm
  ,CbpMCD
  ,CbpEnumerator
;

procedure AddChampsSupreg (TOBPiece : TOB; TOBL : TOB);
begin
	TOBL.AddChampSupValeur ('PIECEORIGINE',EncodeRefPiece (TOBPiece,0,false));
	TOBL.AddChampSupValeur ('GP_DOMAINE',TOBPiece.GetValue('GP_DOMAINE'));
	TOBL.AddChampSupValeur ('GP_REFINTERNE',TOBPiece.GetValue('GP_REFINTERNE'));
  TOBL.AddChampSupValeur ('GP_NUMADRESSEFACT',TOBPiece.GetValue('GP_NUMADRESSEFACT'));
  TOBL.AddChampSupValeur ('GP_NUMADRESSELIVR',TOBPiece.GetValue('GP_NUMADRESSELIVR'));
  TOBL.AddChampSupValeur ('GP_PIECEFRAIS',TOBPiece.GetValue('GP_PIECEFRAIS'));
  TOBL.AddChampSupValeur ('GP_APPLICFGST',TOBPiece.GetValue('GP_APPLICFGST'));
  // ---
  TOBL.AddChampSupValeur ('GP_APPLICFCST',TOBPiece.GetValue('GP_APPLICFCST'));
  TOBL.AddChampSupValeur ('GP_COEFFR',TOBPiece.GetValue('GP_COEFFR'));
  TOBL.AddChampSupValeur ('GP_COEFFC',TOBPiece.GetValue('GP_COEFFC'));
  //
  (*
  TOBL.AddChampSupValeur ('TOTALPA',TOBL.GetValue('GL_DPA')*TOBL.GetValue('GL_QTEFACT'));
  TOBL.AddChampSupValeur ('TOTALPR',TOBL.GetValue('GL_DPR')*TOBL.GetValue('GL_QTEFACT'));
  *)
  TOBL.AddChampSupValeur ('TOTALPA',TOBL.GetValue('GL_MONTANTPA'));
  TOBL.AddChampSupValeur ('TOTALPR',TOBL.GetValue('GL_MONTANTPR'));
  TOBL.AddChampSupValeur ('TOTALPAFG',TOBL.GetValue('GL_MONTANTPAFG'));
  TOBL.AddChampSupValeur ('TOTALPAFC',TOBL.GetValue('GL_MONTANTPAFC'));
  TOBL.AddChampSupValeur ('TOTALPAFR',TOBL.GetValue('GL_MONTANTPAFR'));
  TOBL.AddChampSupValeur ('TOTALFG',TOBL.GetValue('GL_MONTANTFG'));
  TOBL.AddChampSupValeur ('TOTALFC',TOBL.GetValue('GL_MONTANTFC'));
  TOBL.AddChampSupValeur ('TOTALFR',TOBL.GetValue('GL_MONTANTFR'));
  TOBL.AddChampSupValeur ('TOTALPV',TOBL.GetValue('GL_PUHT')*TOBL.GetValue('GL_QTEFACT'));
  TOBL.AddChampSupValeur ('TOTALPVDEV',TOBL.GetValue('GL_PUHTDEV')*TOBL.GetValue('GL_QTEFACT'));
  TOBL.AddChampSupValeur ('TOTALPVTTC',TOBL.GetValue('GL_PUTTC')*TOBL.GetValue('GL_QTEFACT'));
  TOBL.AddChampSupValeur ('TOTALPVTTCDEV',TOBL.GetValue('GL_PUTTCDEV')*TOBL.GetValue('GL_QTEFACT'));
end;

function EcritDebutParagraphe (TOBdevis,TOBCtrEtude,TOBLIG : TOB;Niveau : integer;TheNode : TTreeNode) : TOB;
var variante : boolean;
		UniqueId : string;
begin
	UniqueID := CreateClassID;
	variante := IsVariante (TOBLIG);
	result := NewTOBLigne (TOBCtrEtude,0);
  AddChampsSupreg (TOBDevis,result);
	PieceVersLigne (TOBdevis,result);
  result.cleardetail; // on enleve les filles analytiques
  SetParagraphe (TOBCtrEtude,result,TOBCtrEtude.detail.count-1,niveau,variante);
  result.putValue('Gl_LIBELLE',TheNode.Text);
  //
  result.AddChampSupValeur ('UNIQUEPARAGID',UniqueId);
  //
end;

(* -- Construction d'une LIGNE a partir de LIGNEOUV  -- *)
procedure DefiniTOBL(TOBDevis,TOBCtrEtude : TOB ; Noeud : TTreeNode; TOBL,TOBO : TOB);
var IndZone,Ind,iTableLig : integer;
		NomZone,NomChamp : string;
    TOBRef : TOB;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
	table := Mcd.getTable(mcd.PrefixetoTable('GL'));
  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    NomZone :=(FieldList.Current as IFieldCOM).name;
    ind := Pos ('_',NomZone);
    NomChamp := copy (NomZone,ind+1,255);
    if TOBO.FieldExists ('BLO_'+Nomchamp) then
    begin
    	TOBL.PutValue(NomZone,TOBO.getvalue('BLO_'+NomChamp));
    end;
  end;
  TOBRef := TOBdevis.findFirst(['GL_NUMLIGNE'],[TOBO.getValue('BLO_NUMLIGNE')],True);
  if TOBref <> nil then
  begin
		CopieLigFromRef(TOBL,TOBRef);
    TOBL.PutValue ('GL_IDENTIFIANTWOL',TOBRef.getValue('GL_IDENTIFIANTWOL'));
    TOBL.PutValue ('GLC_FROMBORDEREAU',TOBRef.getValue('GLC_FROMBORDEREAU'));
  end;
  //
  TOBL.PutValue ('BNP_TYPERESSOURCE',TOBO.getValue('BNP_TYPERESSOURCE'));
  TOBL.PutValue ('BNP_LIBELLE',TOBO.getValue('BNP_LIBELLE'));
  //
  TOBL.PutValue ('GLC_NONAPPLICFRAIS',TOBO.getValue('BLO_NONAPPLICFRAIS'));
  TOBL.PutValue ('GLC_NONAPPLICFC',TOBO.getValue('BLO_NONAPPLICFC'));
  TOBL.PutValue ('GLC_NONAPPLICFG',TOBO.getValue('BLO_NONAPPLICFG'));
  //
  TOBL.putValue('GL_REFARTSAISIE',TOBL.getValue('GL_CODEARTICLE'));
  TOBL.putValue('GLC_NATURETRAVAIL',TOBO.getValue('BLO_NATURETRAVAIL'));
  // controle de la quantité
	If TOB(Noeud.data).getValue('QTE') <> TOBL.GetValue('GL_QTEFACT') then
  begin
  	TOBL.PutValue('GL_QTEFACT',TOB(Noeud.data).getValue('QTE'));
  	TOBL.PutValue('GL_QTERESTE',TOB(Noeud.data).getValue('QTE'));
  	TOBL.PutValue('GL_QTESTOCK',TOB(Noeud.data).getValue('QTE'));
    CalculeBaseFraisLigne (TOBL,True);
  end;

  CalculFraisFromLigne (TOBDevis,TOBL,false);

end;

function EcritLigneArticle (TOBInterDOc,TOBdevis,TOBCtrEtude,TOBOuvrage,TOBOuvrageResult,TOBpere,TOBLIG : TOB;TheNode : TTreeNode) : TOB;
var TOBL : TOB;
		TOBI,TOBD,TOBIntL,TOBRecupOuv,TOBDRecupOuv,TOBDet,TOBDETOUV : TOB;
		Niveau,IndiceNomen,Indice : integer;
    Numligne,N1,N2,N3,N4,N5 : integer;
  	MontantAchat,Montantrevient,MontantPV,MontantPVDEV,MontantPVTTC,MontantPVTTCDEV : double;
    UniqueParagId : string;
begin
	if TOBLIG.data = nil then exit;
  TOBDet := TOB(TOB(TheNode.data).Data);
  if TOBDet.NomTable = 'LIGNEOUV' Then
  begin
  	TOBIntL := TOB.Create('LIGNE',nil,-1);
    AddLesSupLigne(TOBIntL, False);
    InitLesSupLigne(TOBIntL);
  	DefiniTOBL(TOBDEVIS,TOBCtrEtude,TheNode,TOBIntL,TOBDet);
    TOBL := TOBIntL;
    //FV1 - 12/04/2017 : FS#2488 - TEAM RESEAUX - Génération contre-études ne reprend pas tout les PU
    if (TOBL.GetValue('GL_TYPELIGNE') = '') AND (TOBL.GetValue('GL_CODEARTICLE') <> '') then TOBL.PutValue('GL_TYPELIGNE', 'ART');
    //
    if (TOBDet.Detail.count > 0) then
    begin
    	Numligne := TOBDet.getValue('BLO_NUMLIGNE');
    	N1 := TOBDet.getValue('BLO_N1');
    	N2 := TOBDet.getValue('BLO_N2');
    	N3 := TOBDet.getValue('BLO_N3');
    	N4 := TOBDet.getValue('BLO_N4');
    	N5 := TOBDet.getValue('BLO_N5');
      TOBDETOUV := TOBOuvrage.findFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],
      																	[Numligne,N1,N2,N3,N4,N5],true);
      if TOBDetOuv <> nil then
      begin
        TOBRecupOuv := TOB.Create ('UN OUVRAGE',TOBOuvrageResult,-1);
        TOBrecupOuv.AddChampSupValeur ('UTILISE','X');

        //TOBDRecupOuv := TOB.Create ('UN OUVRAGE',TOBRecupOuv,-1);
        //TOBDRecupOuv.dupliquer (TOBDETOUV,true,true);
        for Indice := 0 to TOBDetOuv.detail.count -1 do
        begin
          TOBDRecupOuv:=TOB.Create('LIGNEOUV',TOBRecupOuv,-1) ;
          InsertionChampSupOuv (TOBDRecupOuv,false);
        	TOBDRecupOuv.dupliquer (TOBDETOUV.detail[Indice],true,true);
        end;
        TOBL.putValue('GL_INDICENOMEN',TOBOuvrageResult.detail.count);
        //
    		NumeroteLigneOuv (TOBRecupOuv,TOBL,1,1,0,0,0);
        //
      end else
      begin
        TOBL.putValue('GL_INDICENOMEN',0);
      end;
    end;
  end else
  begin
  	TOBL := TOBDet;
    if ((TOBL.GetValue('GL_TYPEARTICLE')='OUV') or (TOBL.GetValue('GL_TYPEARTICLE')='ARP')) and
    		(TOBL.getValue('GL_INDICENOMEN') > 0) then
    begin
    	IndiceNomen := TOBL.getValue('GL_INDICENOMEN');
      if (IndiceNomen >0) and (IndiceNomen-1 <= TOBOUVrage.detail.count) and (TOBOuvrage.detail[IndiceNomen-1] <> nil) then
      begin
        // recup du détail de l'ouvrage associé
        TOBRecupOuv := TOB.Create ('UN OUVRAGE',TOBOuvrageResult,-1);
        TOBRecupOuv.dupliquer (TOBOuvrage.detail[IndiceNomen-1],true,true);
        TOBL.putValue('GL_INDICENOMEN',TOBOuvrageResult.detail.count);
    		NumeroteLigneOuv (TOBRecupOuv,TOBL,1,1,0,0,0);
      end else
      begin
      	TOBL.putValue('GL_INDICENOMEN',0);
      end;
    end else
    begin
      TOBL.putValue('GL_INDICENOMEN',0);
    end;
  end;

  TOBI := TOB.create ('LE LIEN', TOBInterDoc,-1);
  TOBD := TOB.Create ('LIGNE',TOBI,-1);
  TOBD.dupliquer (TOBL,false,true);
  // identificateur du paragraphe
  //
	if TOBPere = nil then
  begin
  	Niveau := 0 ;
  	UniqueParagId := '';
  end else
  begin
		Niveau := TOBPere.getValue('GL_NIVEAUIMBRIC');
    UniqueParagId := TOBpere.GetValue('UNIQUEPARAGID');
  end;
  result := nil;
  if (IsOuvrage (TOBD) and TOBD.GetValue('GL_INDICENOMEN') = 0 ) or (Not ISOuvrage(TOBD)) then
  begin
  	result := TOBCtrEtude.findFirst(['UNIQUEPARAGID','GL_ARTICLE','GL_LIBELLE'],[UniqueParagId,TOBD.GetValue('GL_ARTICLE'),TOBD.GetValue('GL_LIBELLE')],true);
  end;
  if result = nil then
  begin
    result := NewTOBLigne (TOBCtrEtude,0);
    result.dupliquer (TOBL,false,true);
    AddChampsSupReg (TOBDevis,result); // mise en place des champs sup pour le regroupement
    result.AddChampSupValeur ('UNIQUEPARAGID',UniqueParagId);
    result.putValue('GL_NIVEAUIMBRIC',Niveau);
    result.Data := TOBI;
    result.cleardetail;
  end else
  begin
    (*
  	MontantAchat := TOBD.GetValue('GL_DPA')*TOBD.GetValue('GL_QTEFACT');
  	Montantrevient := TOBD.GetValue('GL_DPR')*TOBD.GetValue('GL_QTEFACT');
    *)
  	MontantAchat := TOBD.GetValue('GL_MONTANTPA');
  	Montantrevient := TOBD.GetValue('GL_MONTANTPR');
  	MontantPV := TOBD.GetValue('GL_PUHT')*TOBD.GetValue('GL_QTEFACT');
  	MontantPVDEV := TOBD.GetValue('GL_PUHTDEV')*TOBD.GetValue('GL_QTEFACT');
  	MontantPVTTC := TOBD.GetValue('GL_PUTTC')*TOBD.GetValue('GL_QTEFACT');
  	MontantPVTTCDEV := TOBD.GetValue('GL_PUTTCDEV')*TOBD.GetValue('GL_QTEFACT');
  	// cumul des montants achat revient et vente et calcul de pua,pr,pv
    result.PutValue ('TOTALPA',result.GetValue ('TOTALPA')+MontantAchat);
    result.PutValue ('TOTALPR',result.GetValue ('TOTALPR')+Montantrevient);

    result.PutValue ('TOTALPAFG',result.GetValue ('TOTALPAFG')+TOBL.GetValue('GL_MONTANTPAFG'));
    result.PutValue ('TOTALPAFC',result.GetValue ('TOTALPAFC')+TOBL.GetValue('GL_MONTANTPAFC'));
    result.PutValue ('TOTALPAFR',result.GetValue ('TOTALPAFR')+TOBL.GetValue('GL_MONTANTPAFR'));
    result.PutValue ('TOTALFG',result.GetValue ('TOTALFG')+TOBL.GetValue('GL_MONTANTFG'));
    result.PutValue ('TOTALFC',result.GetValue ('TOTALFC')+TOBL.GetValue('GL_MONTANTFC'));
    result.PutValue ('TOTALFR',result.GetValue ('TOTALFR')+TOBL.GetValue('GL_MONTANTFR'));

    result.PutValue ('TOTALPV',result.GetValue ('TOTALPV')+MontantPV);
    result.PutValue ('TOTALPVDEV',result.GetValue ('TOTALPVDEV')+MontantPVDEV);
    result.PutValue ('TOTALPVTTC',result.GetValue ('TOTALPVTTC')+MontantPVTTC);
    result.PutValue ('TOTALPVTTCDEV',result.GetValue ('TOTALPVTTCDEV')+MontantPVTTCDEV);
    result.putValue ('GL_QTEFACT',result.getValue ('GL_QTEFACT')+TOBD.getValue('GL_QTEFACT'));
    if result.getValue ('GL_QTEFACT') <> 0 then
    begin
    	result.putValue ('GL_DPA',Arrondi(result.getValue ('TOTALPA')/result.getValue('GL_QTEFACT'),V_PGI.okDECP));
    	result.putValue ('GL_DPR',Arrondi(result.getValue ('TOTALPR')/result.getValue('GL_QTEFACT'),V_PGI.okdecp));
    	result.putValue ('GL_MONTANTFG',result.getValue ('TOTALFG'));
    	if result.getValue ('TOTALPA') <> 0 then
      	result.putValue ('GL_COEFFG',result.getValue ('TOTALFG')/result.getValue ('TOTALPA'))
      else
      	result.putValue ('GL_COEFFG',0);
    	result.putValue ('GL_PUHT',Arrondi(result.getValue ('TOTALPV')/result.getValue('GL_QTEFACT'),V_PGI.okdecp));
    	result.putValue ('GL_PUHTDEV',Arrondi(result.getValue ('TOTALPVDEV')/result.getValue('GL_QTEFACT'),V_PGI.okdecp));
    	result.putValue ('GL_PUTTCDEV',arrondi(result.getValue ('TOTALPVTTCDEV')/result.getValue('GL_QTEFACT'),V_PGI.okdecp));
    	result.putValue ('GL_PUTTC',arrondi(result.getValue ('TOTALPVTTC')/result.getValue('GL_QTEFACT'),v_pgi.okdecp));
    end;
  end;
  if TOB(TOBLIG.data).NomTable = 'LIGNEOUV' Then
  begin
  	TOBIntL.free;
  end;
end;

function EcritFinParagraphe (TOBDevis,TOBCtrEtude,TOBDebutPar : TOB;NodeCurr : TTreeNode) : TOB;
var variante : boolean;
begin
	variante := IsVariante (TOBDebutPar);
	result := NewTOBLigne (TOBCtrEtude,0);
  AddChampsSupreg (TOBDevis,result);
	PieceVersLigne (TOBdevis,result);
  result.cleardetail; // on enleve les filles analytiques
  SetParagraphe (TOBCtrEtude,result,TOBCtrEtude.detail.count-1,TOBDebutPar.getValue('GL_NIVEAUIMBRIC'),variante,TtpFin);
  result.putValue('GL_LIBELLE','Total '+NodeCurr.Text);
end;


function EcritTobLigne (TOBInterDoc,TOBdevis,TOBCtrEtude,TOBOuvrage,TOBOuvrageResult,TOBPereLigneCtrEtude,TOBLIG : TOB; NodeCurr : TTreeNode) : TOB;
var Niveau : integer;
		UniqueId : string;
begin
	result := nil;
  if TOBLIG.GetValue('TYPE')='PAR' then
  begin
  	if TOBPereLigneCtrEtude = nil then Niveau := 1
    															else Niveau := TOBPereLigneCtrEtude.getValue('GL_NIVEAUIMBRIC') +1;
  	result := EcritDebutParagraphe (TOBDevis,TOBCtrEtude,TOBLIG,Niveau,NodeCurr);
  end else
  begin
  	result := EcritLigneArticle (TOBInterDoc,TOBDevis,TOBCtrEtude,TOBOuvrage,TOBOuvrageResult,TOBPereLigneCtrEtude,TOBLIG,NodeCurr);
  end;
end;

procedure GereNoeud (TOBInterDOc,TOBDevis,TOBCtrEtude,TOBOuvrage,TOBOuvrageResult : TOB; LignePereCtrEtude : TOB; TheNode : TTreeNode);
var NodeSuiv : TTreeNode;
		CurrLigne , NewLigneCtrEtude : TOB;
    Indice : integer;
begin
  if not assigned (TheNode.data) then exit;
  CurrLigne := TOB(TheNode.data);
  NewLigneCtrEtude := EcritTobLigne (TOBInterDoc,TOBdevis,TOBCtrEtude,TOBOuvrage,TOBOuvrageResult,LignePereCtrEtude,CurrLigne,TheNode);
  if (TheNode.Count > 0) and (CurrLigne.GEtValue('TYPE')='PAR') then
  begin
    for Indice := 0 to TheNode.Count -1 do
    begin
      NodeSuiv := TheNode.Item[Indice];
      GereNoeud (TOBINterDoc,TOBDevis,TOBCtrEtude,TOBOuvrage,TOBOuvrageResult,NewLigneCtrEtude,NodeSuiv);
    end;
  end;
  if CurrLigne.GetValue('TYPE')='PAR' then
  begin
  	EcritFinParagraphe (TOBDevis,TOBCtrEtude,NewLigneCtrEtude,TheNode);
  end;
end;

procedure ArrondiQte (TOBCtrEtude : TOB);
var Indice : integer;
		TOBL : TOB;
begin
  for indice := 0 to TOBCtrEtude.detail.count -1 do
  begin
    TOBL := TOBCtrEtude.detail[Indice];
    TOBL.PutValue('GL_QTEFACT',Arrondi(TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecQ));
  end;
end;

procedure DefinitPieceCtrEtude ( TOBdevis,TOBCtrEtude,TOBOuvrage,TOBOuvrageResult,TOBInterDoc : TOB ; RootCtrEtude : TTreeNode);
var Indice : integer;
		TheNode : TTreeNode;
begin
  TOBCtrEtude.ClearDetail;
  TOBCtrEtude.Dupliquer (TOBdevis,false,true);  // recup de l'entete PIECE
	if RootCtrEtude.Count = 0 then exit;
  for Indice := 0 to RootCtrEtude.Count -1 do
  begin
    TheNode := RootCtrEtude.Item[Indice];
    GereNoeud (TOBInterDoc,TOBdevis,TOBCtrEtude,TOBOuvrage,TOBOuvrageResult,nil,TheNode);
  end;
  ArrondiQte (TOBCtrEtude);
  for indice := 0 to TOBCtrEtude.detail.count -1 do
  begin

  end;
end;

end.
