unit REGENERECONSO_TOF;

interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     fe_main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
	   maineagl,
     eMul,
     uTob,
{$ENDIF}
{$IFDEF BTP}
     CalcOLEGenericBTP,
{$ENDIF}
		 Entgc,
		 Hpanel,
		 Hctrls,
     HStatus,
     forms,
     sysutils,
     ComCtrls,
     HEnt1,
     HTB97,
     HMsgBox,
     UTofAfBaseCodeAffaire,
     UtilSaisieConso,
     UTOF,
     UTOB,uEntCommun,UtilTOBPiece ;

Type
	TModeRech = (TmrFirst,TmrNext);
	TMessErr = (IMErrOrigine,IMErrPrecedent,IMErrArtInconnu,IMErrCptConso,IMErrBLCSansRecep,IMErrEcrAncConso,IMErrEcrNewConso,IMErrEcrInfosSup,IMerrTypeDoc);

  TOF_REGENERECONSO = Class (TOF_AFBASECODEAFFAIRE)
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
  	Affaire : THEdit;
    memo : TMemo;
    TOBLigneVrac,TOBLigneTriee,TOBConsommations,TOBLignesPhases : TOB;
    NbrATraiter : integer;
//    TheWindow : TToolWindow97;
//    Thegraph : THTreeView;
    procedure AddChampsSupRef (TOBLignes : TOB; Fille : boolean);
    function  AddnewConso(TOBCourante, TOBpere, TOBinitiale: TOB): boolean;
    procedure AddReference (TOBLignes : TOB);
    procedure AffecteALaLivraison(NaturePiece: string; TOBPere, TOBL: TOB;NumMouv: double; var QteAffecte, MTPUA, QteAffecteReel: double);
    procedure AffecteALaLivraisonFromArticle(NaturePiece: string; TOBPere,TOBL: TOB; NumMouv: double; var QteAffecte, MTPUA,QteAffecteReel: double);
    function  AffecteLesLivReceptions(TOBpere, TOBL, TOBCL: TOB;NumMouv: double; ReferenceOrigine: string;var PAMoyen: double): Boolean;
    function  AffecteLesReceptions(TOBpere,TOBL,TOBCL: TOB;NumMouv : double;ReferenceOrigine: string; var PAMoyen: double): Boolean;
    procedure AffecteLivraisonsPrealable(TOBCP, TOBC: TOB);
    procedure AppliqueValoLigne(TOBCourante, TOBC: TOB; CoefPaPr,CoefPrPv: double; VenteAchat: string);
    procedure AssocieLigne(Papa, fils: TOB);
  	procedure BLanceTraitementClick (Sender : TObject);
    function CherchesReceptionsAffaireArticle(TOBL, TOBC, TOBPere: TOB;NumMouv: double; var PUA: double): boolean;
    function  ControleSaisie : boolean;
    procedure CreeReference(reference: string);
    procedure CreeTob;
    function  DefiniSequence: boolean;
    function  EnregistreConso(TobCourante, TOBpere, TOBinitiale: TOB) : boolean;
    procedure  EcritlesConso;
    function  IlEstOuPapa(reference: string): TOB;
    function  FindConso(Nummouv: double; indice: integer; Mode : TModeRech=TmrFirst): TOB;
    function  InsereLigne(TOBL: TOB): boolean;
    function  IsPieceANepasPrendreEnCompte (Reference : string) : boolean;
    function  IsReceptionFournisseur(Reference: string): boolean;
    function  IsRegroupement(Piece: string): boolean;
    procedure LanceRegenerationChantier;
    procedure LibereTOB;
    function  MajLignesPhases: boolean;
    procedure MessageErreur(TOBL: TOB; TypeErr: TMessErr);
    procedure NettoieTOB;
    procedure PositionnePiecePrecedente(TOBC, TOBpere: TOB;PiecePrecedente: string; var CoefPaPr : double; var CoefPrPv : double);
    procedure RecupLignes (Affaire : string);
    function RegenereConsommations : boolean;
    procedure RetourChantierPiecePrecedente(TOBC, TobPere: TOB;Pieceprecedente: string);
    procedure TraitementRegeneration ;
    function TraitePere(TOBCourante, TOBParente, TOBinitiale: TOB) : boolean;
  end ;

procedure GestionRegenerationConso;

Implementation

uses factcomm,StockUtil,FactTob;

procedure GestionRegenerationConso;
begin
	AglLanceFiche ('BTP','BTREGENERECONSO','','','ACTION=MODIFICATION');
end;

procedure TOF_REGENERECONSO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_REGENERECONSO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_REGENERECONSO.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_REGENERECONSO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_REGENERECONSO.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolBarButton97(Getcontrol('BLANCETRAITEMENT')).onclick := BLanceTraitementClick;
  Affaire := THEdit (GetCOntrol('AFF_AFFAIRE'));
  memo := TMemo (GetControl('MEMO'));
  Memo.Enabled := false;
  CreeTob;
end ;

procedure TOF_REGENERECONSO.CreeTob;
begin
	TOBLigneVrac := TOB.Create ('LES LIGNES EN VRAC',nil,-1);
  TOBLigneTriee := TOB.create ('LES LIGNES TRIEES',nil,-1);
	TOBConsommations := TOB.Create ('LES CONSO',nil,-1);
  TOBLignesPhases := TOB.Create ('LES LIGNES PHASES',nil,-1);
end;

procedure TOF_REGENERECONSO.OnClose ;
begin
  Inherited ;
	LibereTOB;
end ;

procedure TOF_REGENERECONSO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_REGENERECONSO.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_REGENERECONSO.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin
  Aff0 := THEdit(GetControl('AFF_AFFAIRE0'));
  Aff := THEdit(GetControl('AFF_AFFAIRE'));
  Aff1 := THEdit(GetControl('AFF_AFFAIRE1'));
  Aff2 := THEdit(GetControl('AFF_AFFAIRE2'));
  Aff3 := THEdit(GetControl('AFF_AFFAIRE3'));
  Aff4 := THEdit(GetControl('AFF_AVENANT'));
end;

procedure TOF_REGENERECONSO.BLanceTraitementClick(Sender: TObject);
begin
	NettoieTOB;
	if ControleSaisie then LanceRegenerationChantier;
end;

function TOF_REGENERECONSO.ControleSaisie: boolean;
var Q : TQuery;
begin
  Q := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+Affaire.text+'"',true,-1,'',true);
  result:= not Q.eof;
  if not result then PGIBox (TraduireMemoire('Chantier inconnu'),ecran.caption);
  ferme (Q);
end;

procedure TOF_REGENERECONSO.LanceRegenerationChantier;
begin
 //
 if PgiAsk (TraduireMemoire('Désirez-vous réellement lancer le traitement de regénération des consommations ?'),ecran.caption) = mryes then
 begin
 		TRANSACTIONS (TraitementRegeneration ,1);
 end;
end;

procedure TOF_REGENERECONSO.TraitementRegeneration;
begin
	THPanel(GetControl('PENTETE')).Enabled := false;
	Memo.Enabled := True;
  Memo.Lines.Clear;
  TRY
    RecupLignes (Affaire.text);
    NbrATraiter := TOBLigneVrac.detail.count;
    if not DefiniSequence then BEGIN V_PGI.IOerror := oeUnknown; exit; END;
    if V_PGI.ioError = OeOk then if not RegenereConsommations then exit;
    if V_PGI.ioError = OeOk then TOBConsommations.Detail.Sort ('BCO_NUMMOUV;BCO_INDICE');
    if V_PGI.ioError = OeOk then EcritlesConso;
  FINALLY
  	if V_PGI.ioError = OeOk then PGIInfo (TraduireMemoire('Regénération terminée...'),ecran.caption);
    THPanel(GetControl('PENTETE')).Enabled := True;
    Memo.Enabled := False;
  END;
end;

procedure TOF_REGENERECONSO.RecupLignes(Affaire: string);
var QQ : Tquery;
    Select : string;
begin
	memo.Clear;
  Memo.Lines.Add (TraduireMemoire('1. Récupération des lignes de documents à traiter'));
	//
 	// Phase 1 - recup des Lignes de documents
  //
  Select := 'SELECT LIGNE.*,LIGNEPHASES.* FROM LIGNE LEFT JOIN LIGNEPHASES ON BLP_NATUREPIECEG=GL_NATUREPIECEG AND '+
            'BLP_NUMERO=GL_NUMERO AND BLP_INDICEG=GL_INDICEG AND BLP_NUMORDRE=GL_NUMORDRE '+
  					'WHERE GL_AFFAIRE="'+affaire+'" AND '+
  					'GL_NATUREPIECEG IN ("CF","CFR","BLF","LFR","FF","BFA","AF","LBT","BFC") AND '+
            'GL_TYPELIGNE="ART" ORDER BY GL_PIECEPRECEDENTE,GL_PIECEORIGINE,GL_DATEPIECE';
  QQ := OpenSql (Select,True,-1,'',true);
  if not QQ.eof then
  begin
  	TOBLigneVrac.LoadDetailDB ('LIGNE','','',QQ,false,true);
    AddChampsSupRef (TOBLigneVrac,True);
    AddReference (TOBLigneVrac);
  end;
  ferme (QQ);
end;

function TOF_REGENERECONSO.DefiniSequence : boolean;
var		TOBL : TOB;
begin
	result := true;
  Memo.Lines.Add (TraduireMemoire('3. Définition de l''ordre de lecture'));
  While TOBLIgneVrac.detail.count > 0 do
  begin
  	TOBL := TOBLIgneVrac.detail[0];
    if TOBL <> nil then
    begin
    	if not InsereLigne (TOBL) then
      begin
      	result := false;
        V_PGI.Ioerror := OeUnknown;
      	break;  // pas la peine d'aller plus loin..c foireux
      end;
    end;
  end;
end;

procedure TOF_REGENERECONSO.LibereTOB;
begin
	TOBLigneVrac.free;
  TOBLigneTriee.free;
	TOBConsommations.free;
  TOBLignesPhases.free;
end;

procedure TOF_REGENERECONSO.NettoieTOB;
begin
	TOBLigneVrac.ClearDetail;
  TOBLigneTriee.ClearDetail;
	TOBConsommations.ClearDetail;
  TOBLignesPhases.ClearDetail;
end;

procedure TOF_REGENERECONSO.AddChampsSupRef(TOBLignes: TOB; Fille : boolean);
begin
	TOBLignes.Detail[0].AddChampSup ('REFERENCEPIECE',Fille);
	TOBLignes.Detail[0].AddChampSupValeur ('MAXINDICE',0,Fille);
end;

procedure TOF_REGENERECONSO.AddReference(TOBLignes: TOB);
var Indice : Integer;
		TOBl : TOB;
    RefPiece : string;
begin
  Memo.Lines.Add (TraduireMemoire('2. Phase de préparation initiale'));
  for Indice := 0 to TOBLignes.detail.count -1 do
  begin
    TOBL := TOBLignes.detail[Indice];
    RefPiece := EncodeRefPiece (TOBL);
    TOBL.PutValue('REFERENCEPIECE',RefPiece);
  end;
end;

function TOF_REGENERECONSO.InsereLigne (TOBL : TOB) : boolean;
var TOBP : TOB; // TOB Position parent
begin
	result := true;
	if (TOBL.GetValue('GL_PIECEPRECEDENTE') = '') then
  begin
    if (TOBL.GetValue('GL_PIECEORIGINE') = '') then
    begin
      TOBL.ChangeParent (TOBLigneTriee,-1); // on l'ajoute a la fin
    end else
    begin
  		if IsPieceANepasPrendreEnCompte (TOBL.GetValue('GL_PIECEORIGINE')) then CreeReference(TOBL.GetValue('GL_PIECEORIGINE'));
    	TOBP := TOBLigneTriee.FindFirst(['REFERENCEPIECE'],[TOBL.GetValue('GL_PIECEORIGINE')],true);
      if TOBP = nil then TOBP := IlEstOuPapa (TOBL.GetValue('GL_PIECEORIGINE'));
      if TOBP = nil then
      begin
      	MessageErreur (TOBL,IMErrOrigine);
        result := false;
      end else
      begin
      	AssocieLigne (TOBP,TOBL);
      end;
    end;
  end else
  begin
  	// rattaché a une piece precedente
  	if IsPieceANepasPrendreEnCompte (TOBL.GetValue('GL_PIECEPRECEDENTE')) then CreeReference(TOBL.GetValue('GL_PIECEPRECEDENTE'));
    TOBP := TOBLigneTriee.FindFirst(['REFERENCEPIECE'],[TOBL.GetValue('GL_PIECEPRECEDENTE')],true);
    if TOBP = nil then TOBP := IlEstOuPapa (TOBL.GetValue('GL_PIECEPRECEDENTE'));
    if TOBP = nil then
    begin
      MessageErreur (TOBL,IMErrPrecedent);
      result := false;
    end else
    begin
    	AssocieLigne (TOBP,TOBL);
    end;
  end;
end;

procedure TOF_REGENERECONSO.MessageErreur (TOBL : TOB; TypeErr :TMessErr);
var MessageErr : string;
begin
	if TOBL <> nil then begin
  	MessageErr := traduireMemoire('Erreur sur la ligne ')+IntToStr(TOBL.GetValue('GL_NUMLIGNE')) +TraduireMemoire(' du document ')+InttoStr(TOBL.GetValue('GL_NUMERO'));
  end;
	if TypeErr = IMErrOrigine then MessageErr := MessageErr+ TraduireMemoire(' --> le document d''origine n''existe pas')
  else if TypeErr = IMErrPrecedent then MessageErr := MessageErr+ TraduireMemoire(' --> le document précédent n''existe pas')
  else if TypeErr = IMErrArtInconnu then MessageErr := MessageErr+ TraduireMemoire(' --> l''article n''existe pas')
  else if TypeErr = IMErrCptConso then MessageErr := MessageErr+ TraduireMemoire(' --> Impossible d''attribuer un numéro d''ordre')
  else if TypeErr = IMErrBLCSansRecep then MessageErr := MessageErr+ TraduireMemoire(' --> Une livraison n''a plus de réceptions associées')
  else if TypeErr = IMErrEcrAncConso then MessageErr := MessageErr+ TraduireMemoire(' Erreur durant la suppression des anciennes consommations')
  else if TypeErr = IMErrEcrNewConso then MessageErr := MessageErr+ TraduireMemoire(' Erreur durant l''écriture des nouvelles consommations')
  else if TypeErr = IMErrEcrInfosSup then MessageErr := MessageErr+ TraduireMemoire(' Erreur durant l''écriture des infos. complémentaires consommations')
  else if TypeErr = IMerrTypeDoc then MessageErr := MessageErr+ TraduireMemoire(' --> ce type de document n''est pas géré');

  Memo.Lines.Add (MessageErr);
end;

function TOF_REGENERECONSO.IlEstOuPapa (reference : string) : TOB;
var ThisTOB : TOB;
begin
	result := nil;
  ThisTOB := TOBLIgneVrac.FindFirst(['REFERENCEPIECE'],[reference],true);
  if ThisTOB <> nil then
  begin
  	if not InsereLigne (ThisTOB) then result := nil else result := ThisTOB;
  end;
end;

procedure TOF_REGENERECONSO.AssocieLigne (Papa , fils : TOB);
begin
	Fils.ChangeParent (papa,-1);
end;

function TOF_REGENERECONSO.IsPieceANepasPrendreEnCompte(Reference: string): boolean;
var cledoc : R_CLEDOC;
begin
	result := false;
  DecodeRefPiece (reference,Cledoc);
  if Pos (cledoc.NaturePiece,'PBT;CBT') > 0 then result := true;
end;

procedure TOF_REGENERECONSO.CreeReference(reference : string);
var TOBBidon : TOB;
		cledoc : r_cledoc;
    Q : Tquery;
    Select : string;
begin
	decoderefPiece (reference,Cledoc);
	if TOBLigneTriee.findFirst(['REFERENCEPIECE'],[reference],true) = nil then
  begin
    TOBBidon := TOB.create ('LIGNE',TOBLigneTriee,-1);
    Select := 'SELECT GL_DPR,GL_DPA,GL_PUHT FROM LIGNE WHERE '+WherePiece(Cledoc,TtdLigne,true,true);
    Q := OpenSql (Select,true,-1,'',true);
    if not Q.eof then
    begin
    	TOBBidon.putValue('GL_DPA',Q.findfield('GL_DPA').asFloat);
    	TOBBidon.putValue('GL_DPR',Q.findfield('GL_DPR').asFloat);
    	TOBBidon.putValue('GL_PUHT',Q.findfield('GL_PUHT').asFloat);
    end;
    ferme (Q);
    TOBBidon.AddChampSupValeur ('REFERENCEPIECE',reference);
    TOBBidon.PutValue ('GL_NATUREPIECEG',Cledoc.NaturePiece);
    TOBBidon.PutValue ('GL_SOUCHE',Cledoc.Souche);
    TOBBidon.PutValue ('GL_NUMERO',Cledoc.NumeroPiece);
    TOBBidon.PutValue ('GL_INDICEG',Cledoc.Indice);
    TOBBidon.PutValue ('GL_DATEPIECE',Cledoc.DatePiece);
    TOBBidon.PutValue ('GL_NUMORDRE',Cledoc.NumOrdre);
    TOBBidon.PutValue ('GL_NUMLIGNE',Cledoc.NumLigne);
    TOBBidon.PutValue ('GL_PIECEPRECEDENTE','');
    TOBBidon.PutValue ('GL_PIECEORIGINE','');
  end;
end;

function TOF_REGENERECONSO.TraitePere (TOBCourante,TOBParente,TOBinitiale : TOB) : boolean;
var TOBC : TOB;
		Indice : integer;
begin
	result := true;
	if not EnregistreConso (TobCourante,TOBParente,TOBinitiale) then
  BEGIN
  	result := false;
    V_PGI.IOError := oeUnknown;
    exit;
  END;
  for Indice := 0 to TobCourante.detail.count -1 do
  begin
    TOBC := TobCourante.detail[Indice];
    if not TraitePere (TOBC,TOBCourante,TOBinitiale) then BEGIN result := false; V_PGI.IOError := oeUnknown; break; END;
  end;
end;

function TOF_REGENERECONSO.RegenereConsommations : boolean;
var Indice : integer;
		TOBPAPA : TOB;
begin
	result := true;
  Memo.Lines.Add (TraduireMemoire('4. Génération des consommations'));
	InitMove (NbrATraiter,'Traitement...');
  TRY
    for Indice := 0 to TOBLigneTriee.detail.count -1 do
    begin
      TOBPAPA := TOBLigneTriee.detail[Indice];
      if not TraitePere (TOBPAPA,TOBPAPA,TOBPAPA) then BEGIN result := false;V_PGI.IOError := oeUnknown; break; END;
    end;
  FINALLY
  	FINIMOVE;
  ENd;
end;

function TOF_REGENERECONSO.IsReceptionFournisseur (Reference : string) : boolean;
var Cledoc  : r_cledoc;
begin
	result := false;
  if reference = '' then exit;
	DecodeRefPiece (Reference,cledoc);
  result := (Pos (cledoc.NaturePiece,GetPieceAchat (false,false,false,true)) > 0);
end;

function TOF_REGENERECONSO.IsRegroupement (Piece : string) : boolean;
begin
	Result := (Piece = 'CFR') or (Piece = 'LFR');
end;

procedure TOF_REGENERECONSO.AffecteALaLivraisonFromArticle (NaturePiece: string;TOBPere,TOBL : TOB; NumMouv : double; var QteAffecte : double; var MTPUA : double; var QteAffecteReel : double);
var UneTOB,LaTOBCOnso,TOBNewConso  : TOB;
    QteAPoser : double;
begin

  UneTOB := TOBPere.findFirst(['GL_NATUREPIECEG','GL_ARTICLE'],[NaturePiece,TOBL.GetValue('GL_ARTICLE')],true);
  repeat
  	if UneTOb = nil then break;
    // on a une reception ... on va donc chercher l'element de conso associée
		//LaTOBConso := TOBConsommations.findFirst(['BCO_NUMMOUV','BCO_INDICE'],[UneTOB.GetValue('NEWNUMMOUV'),0],true);
    LaTOBConso := FindConso (UneTOB.GetValue('NEWNUMMOUV'),0);
    if LaTOBconso = nil then exit;
    if (LaTOBConso.getValue('BCO_TRANSFORME') <> 'X') and
    	 (LaTOBConso.GetValue('BCO_QUANTITE') > 0) and
       (LaTOBConso.GetValue('BCO_QUANTITE') - LaTOBConso.GetValue('BCO_QTEVENTE') > 0) and
       (LaTOBConso.getValue('BCO_DATEMOUV') <= TOBL.GetValue('GL_DATEPIECE')) then
    begin
    	// il reste encore une possibilité de livraison sur chantier
      if QteAffecte > (LaTOBConso.GetValue('BCO_QUANTITE')-LaTOBConso.GetValue('BCO_QTEVENTE')) then
      begin
      	QteAPoser := LaTOBConso.GetValue('BCO_QUANTITE') - LaTOBConso.GetValue('BCO_QTEVENTE');
      end else
      begin
      	QteAPoser := QteAffecte;
    	end;
      TOBL.putValue('MAXINDICE',TOBL.GetValue('MAXINDICE')+1);
      LaTOBConso.putValue('BCO_TRAITEVENTE','X');
      LaTOBConso.putValue('BCO_QTEVENTE', LaTOBConso.GetValue('BCO_QTEVENTE') + QteAPoser);

      TOBNewConso := TOB.Create ('CONSOMMATIONS',TOBConsommations,-1);
      TOBNewConso.Dupliquer (LaTOBConso,false,true); // la ligne se rapportant a la livraison de chantier (reception)
      TOBNewConso.putValue('BCO_INDICE',TOBL.GetValue('MAXINDICE'));
      TOBNewConso.putValue('BCO_LIENVENTE',NumMouv);
      TOBNewConso.putValue('BCO_QUANTITE', QteAPoser);
      TOBNewConso.putValue('BCO_TRANSFORME', 'X');
      calculeLaLigne (TOBNewConso);

      MTPUA := MTPUA + TOBNewConso.GetValue('BCO_MONTANTACH');
      QteAffecte := QTeAffecte - QteAPoser;
      QteAffecteReel := QteAffecteReel + QteAPoser;
    end;

    if QteAffecte > 0 then UneTOB := TOBPere.findNext(['GL_NATUREPIECEG','GL_ARTICLE'],[NaturePiece,TOBL.GetValue('GL_ARTICLE')],true)
                      else break;

  until UneTOB = nil;
end;

procedure TOF_REGENERECONSO.RetourChantierPiecePrecedente (TOBC,TobPere : TOB; Pieceprecedente : string);
var TOBCP : TOB;
begin
	TOBCP := FindConso (TOBPere.getValue('NEWNUMMOUV'),0);
  if TOBCP <> nil then
  begin
    TOBC.PutValue('BCO_DPA', TobCP.GetValue('BCO_DPA'));
    TOBC.PutValue('BCO_DPR', TobCP.GetValue('BCO_DPR'));
    TOBC.PutValue('BCO_PUHT',TobCP.GetValue('BCO_PUHT'));
  end;
	TOBC.putValue('BCO_LIENRETOUR',TOBPere.getValue('NEWNUMMOUV'));
end;

procedure TOF_REGENERECONSO.AffecteALaLivraison (NaturePiece: string;TOBPere,TOBL : TOB; NumMouv : double; var QteAffecte : double; var MTPUA : double; var QteAffecteReel : double);
var UneTOB,LaTOBCOnso,TOBNewConso  : TOB;
    QteAPoser : double;
begin

  UneTOB := TOBPere.findFirst(['GL_NATUREPIECEG'],[NaturePiece],true);
  repeat
  	if UneTOb = nil then break;
    // on a une reception ... on va donc chercher l'element de conso associée
		//LaTOBConso := TOBConsommations.findFirst(['BCO_NUMMOUV','BCO_INDICE'],[UneTOB.GetValue('NEWNUMMOUV'),0],true);
    LaTOBConso := FindConso (UneTOB.GetValue('NEWNUMMOUV'),0);
    if LaTOBconso = nil then exit;
    if (LaTOBConso.getValue('BCO_TRANSFORME') <> 'X') and
    	 (LaTOBConso.GetValue('BCO_QUANTITE') > 0) and
       (LaTOBConso.GetValue('BCO_QUANTITE') - LaTOBConso.GetValue('BCO_QTEVENTE') > 0) and
       (LaTOBConso.getValue('BCO_DATEMOUV') <= TOBL.GetValue('GL_DATEPIECE')) then
    begin
    	// il reste encore une possibilité de livraison sur chantier
      if QteAffecte > (LaTOBConso.GetValue('BCO_QUANTITE')-LaTOBConso.GetValue('BCO_QTEVENTE')) then
      begin
      	QteAPoser := LaTOBConso.GetValue('BCO_QUANTITE') - LaTOBConso.GetValue('BCO_QTEVENTE');
      end else
      begin
      	QteAPoser := QteAffecte;
    	end;
      TOBL.putValue('MAXINDICE',TOBL.GetValue('MAXINDICE')+1);
      LaTOBConso.putValue('BCO_TRAITEVENTE','X');
      LaTOBConso.putValue('BCO_QTEVENTE', LaTOBConso.GetValue('BCO_QTEVENTE') + QteAPoser);

      TOBNewConso := TOB.Create ('CONSOMMATIONS',TOBConsommations,-1);
      TOBNewConso.Dupliquer (LaTOBConso,false,true); // la ligne se rapportant a la livraison de chantier (reception)
      TOBNewConso.putValue('BCO_INDICE',TOBL.GetValue('MAXINDICE'));
      TOBNewConso.putValue('BCO_LIENVENTE',NumMouv);
      TOBNewConso.putValue('BCO_QUANTITE', QteAPoser);
      TOBNewConso.putValue('BCO_TRANSFORME', 'X');
      calculeLaLigne (TOBNewConso);

      MTPUA := MTPUA + TOBNewConso.GetValue('BCO_MONTANTACH');
      QteAffecte := QTeAffecte - QteAPoser;
      QteAffecteReel := QteAffecteReel + QteAPoser;
    end;

    if QteAffecte > 0 then UneTOB := TOBPere.findNext(['GL_NATUREPIECEG'],[NaturePiece],true)
                      else break;

  until UneTOB = nil;
end;

function TOF_REGENERECONSO.AffecteLesLivReceptions (TOBpere,TOBL,TOBCL: TOB;NumMouv : double;ReferenceOrigine: string; var PAMoyen: double): Boolean;
var LaTOBR,OneTOB : TOB;
begin
	result := true;
  TOBPere.putValue('MAXINDICE',TOBPere.GetValue('MAXINDICE')+1);
//  LaTOBR := TOBConsommations.findFirst(['BCO_NUMMOUV'],[TOBPere.GetValue('NEWNUMMOUV')],true);
  LaTOBR := FindConso (TOBPere.GetValue('NEWNUMMOUV'),0);
  if LaTOBR = nil then BEGIN result := false; exit; END;
  LaTOBR.putValue('BCO_QTEVENTE', LaTOBR.GetValue('BCO_QUANTITE'));
  LaTOBR.putValue('BCO_TRAITEVENTE','X');
  OneTOB := TOB.create ('CONSOMMATIONS',TOBConsommations,-1);
  OneTOB.Dupliquer (LaTOBR,false,true);
  OneTOB.putValue('BCO_INDICE',TOBPere.GetValue('MAXINDICE'));
  OneTOB.putValue('BCO_LIENVENTE', NumMouv);
  OneTOB.putValue('BCO_TRANSFORME','X');
  PaMoyen := LaTOBR.GetValue('BCO_DPA');
end;

function TOF_REGENERECONSO.CherchesReceptionsAffaireArticle (TOBL,TOBC,TOBPere : TOB;NumMouv : double ; var PUA : double) : boolean;
var QteAffecte,QteAffecteReel : double;
    MTPA : Double;
    QQ : TQuery;
    TOBDispo : TOB;
begin
	result := True;
  QteAffecte := TOBL.GetValue('GL_QTEFACT');
  // recherche des BLF
  AffecteALaLivraisonFromArticle ('BLF',TOBPere,TOBL,NumMouv,QteAffecte,MTPA,QteAffecteReel);
  if QteAffecte > 0 then
  begin
    // recherche des LFR
    AffecteALaLivraisonFromArticle ('LFR',TOBPere,TOBL,NumMouv,QteAffecte,MTPA,QteAffecteReel);
  end;
  if QteAffecte > 0 then
  begin
    // recherche des FF
    AffecteALaLivraisonFromArticle ('FF',TOBPere,TOBL,NumMouv,QteAffecte,MTPA,QteAffecteReel);
  end;
  if QteAffecteReel < QteAffecte then
  begin
  	TOBDispo := TOB.Create ('LES DISPO',nil,-1);
  	// Si la quantité affecte au niveau des recptions est inférieure a la quantité livré alors le reste est pris sur stock.
    TRY
      QQ := OPenSql ('SELECT * FROM DISPO WHERE GQ_ARTICLE="'+
                      TOBL.GetValue('GL_ARTICLE')+'" AND GQ_DEPOT="'+
                      TOBL.GetValue('GL_DEPOT')+'" AND GQ_DATECLOTURE <= "'+
                      USDATETIME (TOBL.getValue('GL_DATEPIECE'))+'" ORDER BY GQ_DATECLOTURE DESC',true,-1,'',true);
      if not QQ.eof then
      begin
      	TOBDispo.loadDetailDB ('DISPO','','',QQ,false);
        MTPA := MTPA + ((QteAffecte-QteAffecteReel) * TOBDispo.detail[0].getValue('GQ_PMAP'));
      end;
      ferme(QQ);
    FINALLY
    	TOBDispo.free;
    End;
  end;
  PUA := Arrondi(MTPA /TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecP);
end;

function TOF_REGENERECONSO.AffecteLesReceptions(TOBpere,TOBL,TOBCL: TOB;NumMouv : double;ReferenceOrigine: string; var PAMoyen: double): Boolean;
var QteAffecte,QteAffecteReel : double;
    MTPA : Double;
    QQ : TQuery;
    TOBDispo : TOB;
begin
	result := True;
  QteAffecte := TOBL.GetValue('GL_QTEFACT');
  // recherche des BLF
  AffecteALaLivraison ('BLF',TOBPere,TOBL,NumMouv,QteAffecte,MTPA,QteAffecteReel);
  if QteAffecte > 0 then
  begin
    // recherche des LFR
    AffecteALaLivraison ('LFR',TOBPere,TOBL,NumMouv,QteAffecte,MTPA,QteAffecteReel);
  end;
  if QteAffecte > 0 then
  begin
    // recherche des FF
    AffecteALaLivraison ('FF',TOBPere,TOBL,NumMouv,QteAffecte,MTPA,QteAffecteReel);
  end;
  if QteAffecteReel < QteAffecte then
  begin
  	TOBDispo := TOB.Create ('LES DISPO',nil,-1);
  	// Si la quantité affecte au niveau des recptions est inférieure a la quantité livré alors le reste est pris sur stock.
    TRY
      QQ := OPenSql ('SELECT * FROM DISPO WHERE GQ_ARTICLE="'+
                      TOBL.GetValue('GL_ARTICLE')+'" AND GQ_DEPOT="'+
                      TOBL.GetValue('GL_DEPOT')+'" AND GQ_DATECLOTURE <= "'+
                      USDATETIME (TOBL.getValue('GL_DATEPIECE'))+'" ORDER BY GQ_DATECLOTURE DESC',true,-1,'',true);
      if not QQ.eof then
      begin
      	TOBDispo.loadDetailDB ('DISPO','','',QQ,false);
        MTPA := MTPA + ((QteAffecte-QteAffecteReel) * TOBDispo.detail[0].getValue('GQ_PMAP'));
      end;
      ferme(QQ);
    FINALLY
    	TOBDispo.free;
    End;
  end;
  PAMoyen := Arrondi(MTPA /TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecP);
end;

procedure TOF_REGENERECONSO.AffecteLivraisonsPrealable (TOBCP,TOBC : TOB);

	function GetConso (NumConso : double; NIndice : integer = -1) : TOB;
  var Indice : integer;
  begin
  	result := nil;
		for Indice := 0 to TOBconsommations.detail.count -1 do
    begin
    	if TOBconsommations.detail[Indice].GetValue ('BCO_NUMMOUV')=NumConso then
      begin
      	if (NIndice = -1) or (TOBconsommations.detail[Indice].GetValue ('BCO_INDICE')=NIndice) then
        begin
      		result := TOBconsommations.detail[Indice];
        	break;
        end;
      end;
    end;
  end;

var TOBCGP,TOBCVENTE : TOB; // TOBC conso grand-pere
		CoefPaPr : double;
    Indice : integer;
begin
	Indice := 0;
  repeat
    TOBCGP := GetConso (TOBCP.getValue('BCO_LIENTRANSFORME'),Indice);
    if TOBCGP <> nil then
    begin
      if TOBCGP.GetValue('BCO_LIENVENTE')<> 0 then
      begin
        TOBCVENTE := GetConso (TOBCGP.getValue('BCO_LIENVENTE'));
        if TOBCVENTE <> nil then
        begin
          if TOBCVENTE.GetValue('BCO_DPA') <> 0 then CoefPapr := TOBCVENTE.GetValue('BCO_DPR') / TOBCVENTE.GetValue('BCO_DPA')
                                                else CoefPapr := 1;
          TOBCVENTE.PutValue('BCO_DPA',TOBC.GetValue('BCO_DPA'));
          TOBCVENTE.PutValue('BCO_DPR',TOBC.GetValue('BCO_DPA')*CoefPapr);
        end;
      end;
      // on continue dans le parcours
      if (TOBCGP.getValue('BCO_LIENTRANSFORME')<>0) then AffecteLivraisonsPrealable (TOBCGP,TOBC);
      inc (Indice);
    end;
  until TOBCGP = nil;
end;



procedure TOF_REGENERECONSO.PositionnePiecePrecedente (TOBC, TOBpere: TOB;PiecePrecedente: string; var CoefPaPr : double; var CoefPrPv : double);
var TOBCP : TOB;
begin
//	TOBCP := TOBConsommations.FindFirst (['BCO_NUMMOUV','BCO_INDICE'],[TOBPere.GetValue('NEWNUMMOUV'),0],true);
  TOBCP := FindConso (TOBPere.GetValue('NEWNUMMOUV'),0);
  if TOBCP <> nil then
  begin
  	TOBCP.putValue('BCO_QUANTITE',TOBCP.GEtVAlue('BCO_QUANTITE')-TOBC.GetValue('BCO_QUANTITE'));
    if TOBCP.GetValue('BCO_LIENTRANSFORME')<>0 then AffecteLivraisonsPrealable (TOBCP,TOBC);
    if TOBCP.GEtVAlue('BCO_QUANTITE') < 0 then TOBCP.putValue('BCO_QUANTITE',0);
    if TOBCP.GetValue('BCO_QUANTITE') = 0 then TOBCP.putValue('BCO_TRANSFORME','X');
    if TOBCP.GetValue('BCO_DPA') <> 0 then CoefPaPR := TOBCP.GetValue('BCO_DPR')/TOBCP.GetValue('BCO_DPA')
    																	else CoefPaPr := 1;
    if TOBCP.GetValue('BCO_DPR') <> 0 then CoefPRPV := TOBCP.GetValue('BCO_PUHT')/TOBCP.GetValue('BCO_DPR')
    																	else CoefPrPv := 1;

    TOBC.PutValue('BCO_LIENTRANSFORME',TOBPere.GetValue('NEWNUMMOUV'));
    calculeLaLigne (TOBCP);
  end;
end;

procedure TOF_REGENERECONSO.AppliqueValoLigne (TOBCourante,TOBC : TOB;CoefPaPr,CoefPrPv : double; VenteAchat : string);
begin
  if (venteAchat = 'ACH') then
  begin
    TOBC.PutValue('BCO_DPA', TOBCourante.GetValue('GL_PUHTNET')/(TOBCourante.getValue('GL_PRIXPOURQTE')));
    TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*CoefPaPr,V_PGI.okdecP));
    TOBC.PutValue('BCO_PUHT',Arrondi(TOBC.GEtValue('BCO_DPR')*CoefPRPV,V_PGI.okdecP));
  end else
  begin
    TOBC.PutValue('BCO_DPA', (TOBCourante.GetValue('GL_DPA')/(TOBCourante.getValue('GL_PRIXPOURQTE'))));
    TOBC.PutValue('BCO_DPR', (TOBCourante.GetValue('GL_DPA')/(TOBCourante.getValue('GL_PRIXPOURQTE'))*CoefPaPr));
    TOBC.PutValue('BCO_PUHT', (TOBCourante.GetValue('GL_DPR')/(TOBCourante.getValue('GL_PRIXPOURQTE'))*CoefPrPv));
  end;
end;

function TOF_REGENERECONSO.AddnewConso (TOBCourante,TOBpere,TOBinitiale : TOB) : boolean;
Var Part,Part0,Part1,Part2,Part3,Part4,Requete,Nature,VenteAchat,TypeArticle,NaturePiece : string;
		PieceOrigine,PiecePrecedente : string;
    Annee,Mois,Jour      : word;
    Semaine : integer;
    DateMouv  : TDateTime;
    Q         : TQuery;
    TheRetour : TGncERROR;
    TOBC,TOBCP : TOB;
    RatioVente,NumMouv,CoefPAPR,COEfPrPV,PUA : double;
Begin
  result := true;
  if TOBCourante = nil then Exit;

  if TOBCourante.GetValue('GL_AFFAIRE') = '' then exit;
  //
  MoveCur(False);            // pour faire avancer le SHIMILIMILI
  //
  PieceOrigine := TOBCourante.GetValue('GL_PIECEORIGINE');
  PiecePrecedente := TOBCourante.GetValue('GL_PIECEPRECEDENTE');
  //
  VenteAchat := GetInfoParPiece(TOBCourante.GetValue('GL_NATUREPIECEG'), 'GPP_VENTEACHAT');
  NaturePiece := TOBCourante.getValue('GL_NATUREPIECEG');
  Part  := TOBCourante.GetValue('GL_AFFAIRE');
  Part0 := '';
  Part1 := '';
  Part2 := '';
  Part3 := '';
  Part4 := TOBCourante.GetValue('GL_AVENANT');

  DateMouv := TOBCourante.GetValue('GL_DATEPIECE');

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
  TypeArticle := TOBCourante.GetValue('GL_TYPEARTICLE');
  // Recherche de la Nature du Mouvement si Prestation
  if TypeArticle = 'PRE' then
  Begin
    Requete := 'SELECT GA_ARTICLE,GA_LIBELLE,N1.BNP_TYPERESSOURCE,N1.BNP_LIBELLE FROM ARTICLE ART ' +
               'LEFT JOIN NATUREPREST N1 ON ART.GA_NATUREPRES=N1.BNP_NATUREPRES ' +
               'WHERE GA_TYPEARTICLE="PRE" AND GA_ARTICLE="'+TOBCourante.GetValue('GL_ARTICLE')+'"';
    Q := OpenSQL(Requete, True,-1,'',true);
    if not Q.eof then
    begin
      Nature := Q.findfield('BNP_TYPERESSOURCE').asString ;
      ferme (Q);
    end else
    Begin
      ferme (Q);
      MessageErreur (TOBCourante,IMErrArtInconnu);
      result := false;
      exit;
    end;
  end;

  if venteAchat = 'ACH' then
  begin
  	CoefPaPr := 0;
    CoefPrPv := 0;
  	RatioVente := GetRatio(TOBCourante, nil, trsVente);
		RecupCoefs (TOBCourante,VenteAchat,CoefPaPr,CoefPrPv);
  end else
  begin
  	RatioVente := 1;
  	CoefPaPr := 0;
    CoefPrPv := 0;
    if TOBCourante.GetValue('GL_DPA') <> 0 then CoefPaPR := TOBCourante.GetValue('GL_DPR')/TOBCourante.GetValue('GL_DPA');
    if TOBCourante.GetValue('GL_DPR') <> 0 then CoefPRPV := TOBCourante.GetValue('GL_PUHT')/TOBCourante.GetValue('GL_DPR');
	end;

  // Calcul du Numéro de Mouvement
  TheRetour := GetNumUniqueConso (NumMouv);
  if TheRetour = gncAbort then
  BEGIN
  	result := false;
    MessageErreur (TOBCourante,IMErrCptConso);
    V_PGI.ioError := OeUnknown;
    Exit;
  END;

  // Chargement de la TOB Conssomation

  TOBC := Tob.Create('CONSOMMATIONS', TobConsommations, -1);

  TOBC.PutValue('BCO_AFFAIRE', TOBCourante.GetValue('GL_AFFAIRE'));
  TOBC.PutValue('BCO_AFFAIRE0', Part0);
  TOBC.PutValue('BCO_AFFAIRE1', Part1);
  TOBC.PutValue('BCO_AFFAIRE2', Part2);
  TOBC.PutValue('BCO_AFFAIRE3', Part3);
  TOBC.PutValue('BCO_PHASETRA', TOBCourante.getValue('BLP_PHASETRA'));
  TOBC.PutValue('BCO_MOIS', Mois);
  TOBC.PutValue('BCO_SEMAINE', Semaine);
  TOBC.PutValue('BCO_DATEMOUV', DateMouv);
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

  TOBC.PutValue('BCO_LIBELLE', TOBCourante.GetValue('GL_LIBELLE'));
  TOBC.PutValue('BCO_TRANSFORME', '-');
  TOBC.PutValue('BCO_NUMMOUV', NumMouv);
  TOBC.PutValue('BCO_INDICE',0);

  if TOBCourante.getValue('GL_PRIXPOURQTE')= 0 then TOBCourante.putvalue('GL_PRIXPOURQTE',1);
  TOBC.PutValue('BCO_CODEARTICLE', TOBCourante.GetValue('GL_CODEARTICLE'));
  TOBC.PutValue('BCO_RESSOURCE', TOBCourante.GetValue('GL_RESSOURCE'));
  TOBC.PutValue('BCO_ARTICLE', TOBCourante.GetValue('GL_ARTICLE'));
  TOBC.PutValue('BCO_QUANTITE', TOBCourante.GetValue('GL_QTEFACT') * RatioVente); // on ne peut pas se baser sur le QTERESTE
  TOBC.PutValue('BCO_QUALIFQTEMOUV', TOBCourante.GetValue('GL_QUALIFQTEVTE'));
  TOBCourante.AddChampSupValeur ('NEWNUMMOUV',Nummouv);  // Stockage du nouveau numero de mouvement dans la ligne de doc
  if (NaturePiece = 'LBT') and not (IsReceptionFournisseur(PieceOrigine)) then
  begin
  	// cas d'une livraison provenant d'un besoin de chantier
    // Affectation de la quantité livrée au reception ou facture achat active.
    if PieceOrigine <> '' then
    begin
    	PUA := TOBCourante.getValue('GL_DPA');
      if not AffecteLesReceptions (TOBinitiale,TOBCourante,TOBC,Nummouv,PieceOrigine,PUA) then
      begin
        MessageErreur (TOBCourante,IMErrBLCSansRecep);
        result := false;
    		V_PGI.ioError := OeUnknown;
        exit;
      end;
      TOBC.putValue('BCO_DPA',PUA);
      RecupCoefs (TOBCourante,venteachat,CoefPaPr,CoefPrPv);
      TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*CoefPaPr,V_PGI.okdecP));
      TOBC.PutValue('BCO_PUHT',TOBCourante.getValue('GL_PUHT')); // pas touche au PV
    end else
    begin
    	// cas de livraisons sans besoins de chantiers, ni livraisons directes depuis une reception ou facture
    	if not  CherchesReceptionsAffaireArticle (TOBCourante,TOBC,TOBLigneTriee,Nummouv,PUA) then
      begin
        MessageErreur (TOBCourante,IMErrBLCSansRecep);
				result := false;
        V_PGI.ioError := OeUnknown;
        exit;
      end;
      TOBC.putValue('BCO_DPA',PUA);
    	AppliqueValoLigne (TOBCourante,TOBC,CoefPaPr,CoefPrPv, VenteAchat);
    end;
//    TOBC.PutValue('BCO_TRANSFORME','X');
  end else if (NaturePiece = 'LBT') and (IsReceptionFournisseur(TOBCourante.getValue('GL_PIECEORIGINE'))) then
  begin
    // livraison chantier issue d'une réception
    if not AffecteLesLivReceptions (TOBPere,TOBCourante,TOBC,Nummouv,PieceOrigine,PUA) then
    begin
      MessageErreur (TOBCourante,IMErrBLCSansRecep);
      result := false;
    	V_PGI.ioError := OeUnknown;
      exit;
    end;
    TOBC.putValue('BCO_DPA',PUA);
    TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*CoefPaPr,V_PGI.okdecP));
    TOBC.PutValue('BCO_PUHT',Arrondi(TOBC.GEtValue('BCO_DPR')*CoefPRPV,V_PGI.okdecP));
  end else if (IsRegroupement(NaturePiece)) then
  begin
  	// gestion DES CFR..LFR
  	// regroupement de commande ou de reception
    TOBCP := FindConso (TOBPere.GetValue('NEWNUMMOUV'),0);
    if TOBCP <> nil then
    begin
      TOBCP.PutValue('BCO_TRANSFORME','X');
      TOBCP.PutValue('BCO_QUANTITE',0);
      calculeLaLigne (TOBCP);
      TOBC.putValue('BCO_LIENTRANSFORME',TOBPere.GetValue('NEWNUMMOUV'));
      TOBC.putValue('BCO_TRAITEVENTE',TOBCP.GetValue('BCO_TRAITEVENTE'));
      TOBC.putValue('BCO_QTEVENTE',TOBCP.GetValue('BCO_QTEVENTE'));
      TOBC.putValue('BCO_DPA',TOBCP.GetValue('BCO_DPA'));
      TOBC.putValue('BCO_DPR',TOBCP.GetValue('BCO_DPR'));
      TOBC.putValue('BCO_PUHT',TOBCP.GetValue('BCO_PUHT'));
    end;
  end else if ( Pos (NaturePiece,GetPieceAchat(false,false,false,True)) > 0) then
  begin
  	// reception ou facture
    TOBC.PutValue('BCO_DPA', TobCourante.GetValue('GL_PUHTNET')/(TobCourante.getValue('GL_PRIXPOURQTE')));
    TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*CoefPaPr,V_PGI.okdecP));
    TOBC.PutValue('BCO_PUHT',Arrondi(TOBC.GEtValue('BCO_DPR')*CoefPRPV,V_PGI.okdecP));
    if PiecePrecedente <> '' then
    begin
    	PositionnePiecePrecedente (TOBC,TobPere,Pieceprecedente,CoefPaPr,CoefPrPv);
    end else
    begin
    	if PieceOrigine = '' then TOBC.PutValue('BCO_QTEINIT', TOBC.GetValue('BCO_QUANTITE'));
    end;
  end else if ( Pos (NaturePiece,GetPieceAchat (False,true,false,false)) > 0 ) and (Pieceorigine='') then
  begin
  	// CF
    TOBC.PutValue('BCO_DPA', TobCourante.GetValue('GL_PUHTNET')/(TobCourante.getValue('GL_PRIXPOURQTE')));
    TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*CoefPaPr,V_PGI.okdecP));
    TOBC.PutValue('BCO_PUHT',Arrondi(TOBC.GEtValue('BCO_DPR')*CoefPRPV,V_PGI.okdecP));
    TOBC.PutValue('BCO_QTEINIT', TOBC.GetValue('BCO_QUANTITE'));
  end else if (Naturepiece = 'BFC') or (Naturepiece = 'BFA') then
  begin
    TOBC.PutValue('BCO_DPA', TobCourante.GetValue('GL_DPA'));
    TOBC.PutValue('BCO_DPR', TobCourante.GetValue('GL_DPR'));
    TOBC.PutValue('BCO_PUHT',TobCourante.GetValue('GL_PUHT'));
  	if PiecePrecedente <> '' then
    begin
    	RetourChantierPiecePrecedente (TOBC,TobPere,Pieceprecedente);
    end;
  end else if (NaturePiece = 'AF') or (NaturePiece = 'AFS') then
  begin
    TOBC.PutValue('BCO_DPA', TobCourante.GetValue('GL_PUHTNET')/(TobCourante.getValue('GL_PRIXPOURQTE')));
    TOBC.PutValue('BCO_DPR', Arrondi(TOBC.GEtValue('BCO_DPA')*CoefPaPr,V_PGI.okdecP));
    TOBC.PutValue('BCO_PUHT',Arrondi(TOBC.GEtValue('BCO_DPR')*CoefPRPV,V_PGI.okdecP));
  end else
  begin
  	MessageErreur (TOBCourante,IMerrTypeDoc);
    result := false;
    V_PGI.IoError := OeUnknown;
    exit;
  end;

  TOBC.PutValue('BCO_FACTURABLE', 'N');
  TOBC.PutValue('BCO_VALIDE', '-');
  TOBC.PutValue('BCO_NATUREPIECEG', TOBCourante.GetValue('GL_NATUREPIECEG'));
  TOBC.PutValue('BCO_SOUCHE', TOBCourante.GetValue('GL_SOUCHE'));
  TOBC.PutValue('BCO_NUMERO', TOBCourante.GetValue('GL_NUMERO'));
  TOBC.PutValue('BCO_INDICEG', TOBCourante.GetValue('GL_INDICEG'));
  TOBC.PutValue('BCO_NUMORDRE', TOBCourante.GetValue('GL_NUMORDRE'));
  CalculeLaLigne(TOBC);
end;

function TOF_REGENERECONSO.EnregistreConso (TobCourante,TOBpere,TOBinitiale : TOB) : boolean;
begin
	result := true;
  if IsPieceANepasPrendreEnCompte (TobCourante.GetValue('REFERENCEPIECE')) then exit;
  result := AddNewConso (TOBCourante,TOBPere,TOBinitiale);
  if not result then V_PGI.Ioerror := oeUnknown;
end;

function TOF_REGENERECONSO.FindConso(Nummouv: double; indice: integer; Mode : TModeRech=TmrFirst): TOB;
begin
	if Indice = -1 then
  begin
    if Mode = TmrFirst then result := TOBConsommations.findFirst(['BCO_NUMMOUV'],[Nummouv],true)
                       else result := TOBConsommations.findNext (['BCO_NUMMOUV'],[Nummouv],true);
  end else
  begin
    if Mode = TmrFirst then result := TOBConsommations.findFirst(['BCO_NUMMOUV','BCO_INDICE'],[Nummouv,Indice],true)
                       else result := TOBConsommations.findNext (['BCO_NUMMOUV','BCO_INDICE'],[Nummouv,Indice],true);
  end;
end;

procedure TOF_REGENERECONSO.EcritlesConso;
var Select : string;
begin
  Memo.Lines.Add (TraduireMemoire('5. Ecriture des consommations'));

	// 1ere phase on supprime les anciennes conso
  Select := 'DELETE FROM CONSOMMATIONS WHERE BCO_AFFAIRE="'+Affaire.text+'" AND BCO_NATUREPIECEG<>"" ';
  if ExecuteSQL (Select) < 0 then
  BEGIN
  	MessageErreur (nil,IMErrEcrAncConso);
		V_PGI.IoError := OeUnknown;
  end;
  // 2eme phase ecriture des nouvelles consommations
  if V_PGI.IoError = OeOk then if not TOBConsommations.InsertDB (nil,true) then
  BEGIN
  	MessageErreur (nil,IMErrEcrNewConso);
  	V_PGI.IoError := OeUnknown;
  END;
  // 3eme phase ecriture des nouvelles consommations
  if V_PGI.IoError = OeOk then if not MajLignesPhases then
  BEGIN
  	MessageErreur (nil,IMErrEcrInfosSup);
  	V_PGI.IoError := OeUnknown;
  END;
end;

function TOF_REGENERECONSO.MajLignesPhases : boolean;

	Function  TraiteLignePhases (TOBLP : TOB) : boolean;
  var indice : integer;
  		TOBLLP,TOBNEW : TOB;
  		Select : string;
  begin
  	result := true;
  	if TOBLP.GetValue('NEWNUMMOUV') <> 0 then
    begin
    	if not ExisteSQL ('SELECT BLP_NUMMOUV FROM LIGNEPHASES '+
      					'WHERE BLP_NATUREPIECEG="'+TOBLP.GetValue('GL_NATUREPIECEG')+'" AND '+
                'BLP_NUMERO='+INtTOStr(TOBLP.GetValue('GL_NUMERO'))+' AND '+
                'BLP_INDICEG='+INtTOStr(TOBLP.GetValue('GL_INDICEG'))+' AND '+
                'BLP_NUMORDRE='+INtTOStr(TOBLP.GetValue('GL_NUMORDRE'))) then
      begin
         TOBNEW := TOB.Create ('LIGNEPHASES',nil,-1);
         TRY
           TOBNEw.PutValue('BLP_NATUREPIECEG',TOBLP.GetValue('GL_NATUREPIECEG'));
           TOBNEw.PutValue('BLP_NUMERO',TOBLP.GetValue('GL_NUMERO'));
           TOBNEw.PutValue('BLP_INDICEG',TOBLP.GetValue('GL_INDICEG'));
           TOBNEw.PutValue('BLP_NUMORDRE',TOBLP.GetValue('GL_NUMORDRE'));
           TOBNEw.PutValue('BLP_NUMMOUV',TOBLP.GetValue('NEWNUMMOUV'));
         if not TOBNew.InsertDB (nil,false) then V_PGI.IOError := OeUnknown;;
         FINALLY
         	TOBNew.free;
         END;
      end else
      begin
        Select := 'UPDATE LIGNEPHASES SET BLP_NUMMOUV='+ floatToStr(TOBLP.GetValue('NEWNUMMOUV'))+' '+
                  'WHERE BLP_NATUREPIECEG="'+TOBLP.GetValue('BLP_NATUREPIECEG')+'" AND '+
                  'BLP_NUMERO='+INtTOStr(TOBLP.GetValue('BLP_NUMERO'))+' AND '+
                  'BLP_INDICEG='+INtTOStr(TOBLP.GetValue('BLP_INDICEG'))+' AND '+
                  'BLP_NUMORDRE='+INtTOStr(TOBLP.GetValue('BLP_NUMORDRE'));
      	if ExecuteSQL (Select) < 0 then BEGIN result := false; V_PGI.IOError := OeUnknown; end;
      end;
    end;
    if TOBLP.Detail.count > 0 then
    begin
    	for Indice := 0 to TOBLP.detail.count -1 do
      begin
      	TOBLLP := TOBLP.detail[Indice];
    		TraiteLignePhases (TOBLLP);
      end;
    end;
  end;

var TOBLP : TOB;
		Indice : integer;
begin
	result := true;
  for Indice := 0 to TOBLigneTriee.detail.count -1 do
  begin
  	TOBLP := TOBLigneTriee.detail[Indice];
    if not TraiteLignePhases (TOBLP) then begin result := false; V_PGI.IOError := OeUnknown; end;
  end;
end;

Initialization
  registerclasses ( [ TOF_REGENERECONSO ] ) ;
end.

