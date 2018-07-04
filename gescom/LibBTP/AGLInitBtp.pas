unit AGLInitBtp;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, ExtCtrls, HTB97, StdCtrls, HPanel, UIUtil, Hent1, Menus,
  HSysMenu, Mask, Buttons,
  {$IFNDEF EAGLCLIENT}
  Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  DBCtrls,Hdb,Mul,fe_main,edtRetat,
  {$ELSE}
  eMul,MainEagl,HQry,UtilEagl,
  {$ENDIF}
  HStatus, hmsgbox,UTOF, UtilPGI, UTOM,
  UTOB, HFLabel, Ent1, SaisUtil, LookUp, Math, FactUtil, EntGC,FactSpec,
  FactCalc, StockUtil,M3FP,  AglInit, FactComm, FactCpta, Facture,
  AdressePiece, Clipbrd, AffaireUtil, ComCtrls, HRichEdt, HRichOLE,
  FactNomen, VentAna, SaisComm, Doc_Parser, LigNomen,
  DimUtil, LigDispoLot, UTofGCPieceArtLie,ShellAPI, DicoBTP, FactGrp,
  BTPUtil, Etudes,EtudesUtil,UtilArticle,UtofBTChangeCodeArt,ParamSoc,FactAcompte,FactRg,
  UTOFBTClotureDev,BTPLANNIFCH_TOF,FactOuvrage,UPlannifchUtil,FactTOB,BTGENCONTRETU_TOF,
  Utof_VideInside
	,FactTvaMilliem,uEntCommun,UTOFBtAnalDev,UDemandePrix
  ;

const
	// libellés des messages
	TexteMessage: array[1..13] of string 	= (
          {1}        'Confirmez-vous l''acceptation des devis sélectionnés ?'
          {2}       ,'Confirmez-vous le refus des devis sélectionnés ?'
          {3}       ,'Confirmez-vous la réactivation des devis sélectionnés ?'
          {4}       ,'Confirmez-vous la facturation des devis sélectionnés ? '
          {5}       ,' sur l''affaire '
          {6}       ,'Le transfert des lignes d''activité a échoué, voulez-vous annuler le refus ?'
          {7}       ,'Confirmez-vous l''acceptation des appels d''offre sélectionnés?'
          {8}       ,'Confirmez-vous le changement de codification des articles sélectionnés ?'
          {9}       ,'Confirmez-vous le traitement de clôture des chantiers sélectionnés ?'
          {10}      ,'Confirmez-vous la validation des études sélectionnées ?'
          {11}      ,'Confirmez-vous l''annulation de l''acceptation des devis sélectionnés ?'
          {12}      ,'Confirmez-vous le lancement de l''analyse des documents sélectionnés ?'
          {13}      ,'Veuillez sélectionner le (ou les) document(s) à analyser.'
                       );

type
     TModegestion = (TmgIntervenant,TmgCotraitance,TmgSousTraitance);
     //
     TAcceptationDocument = class
     public
     //
       TOBPiece       : TOB;
       TOBTiers       : TOB;
       TOBAcomptes    : TOB;
       TOBAcomptes_O  : TOB;
       TobPieceRG     : TOB;
       TOBBasesRg     : TOB;
       TOBPieceTrait  : TOB;
       TOBPieceTrait_O  : TOB;
       //
       QQ             : Tquery;
       cledoc         : R_Cledoc;
       Sql            : string;
       NaturePiece    : String;
       Souche         : String;
       Tiers          : string ;
       Numero         : Integer;
       Indice         : integer;
       Result         : Boolean;
       pass           : Boolean;
       ReglePiece     : Boolean;
       Fgestion       : Tmodegestion;
     //
     private
       fecrModifiable : Boolean;
    	 indiceOuv : Integer;
       // Reglement comptabilisés -----
       TOBPiece_O : TOB;
       TOBOuvrages    : TOB;
       TOBOuvragesP 	: TOB;
       TOBBases       : TOB;
       TOBBasesL      : TOB;
       TOBEches       : TOB;
       TOBAffaireInterv : TOB;
       TOBArticles    : TOB;
       TOBPorcs       : TOB;
       TOBAnaP        : TOB;
       TOBAnaS        : TOB;
       TOBSSTrait     : TOB;
       TOBCPTA        : TOB;
       DEV : RDevise;
       OldEcr, OldStk	: RMVT;
       // ------------------------------

       procedure LibereTOBS;
       procedure CreateTobs;
       procedure ValideLaPieceAcompte;
       procedure LoadLesPieceTrait;

     public
       destructor Destroy; override;
       constructor Create;
       Procedure ChargeLesTObs;
       Procedure DemandeAcompte;
       procedure GereReglements;
       function EcriturePieceModifiable (TOBpiece : TOB) : Boolean;
     end;

  procedure AglPlannificationChantier (Parms : array of variant ; nb : integer) ;
  procedure AglGenereContreEtude (Parms : array of variant ; nb : integer) ;
  Function AGLCreerPieceBTP( parms: array of variant; nb: integer ): variant;

  procedure SaisieAvancementChantier (NaturePiece, Datepiece, Souche, Tiers, AffaireRef : String; NumeroPiece, Indice : integer; Action : TActionFiche ) ;
  Procedure DecisionStockReappro;

  Function DemandeAcompteOk (NaturePiece,Souche,Tiers : string ; Numero,Indice: integer; pass : boolean=false): boolean;
  procedure AglAvancementChantier (Parms : array of variant ; nb : integer) ;
  Procedure PositionneEtatAffaire(CodeAffaire, CodeEtat : String);
  procedure PleinEcran (TheForm : TForm) ;

implementation
uses factvariante,
     BTPrepaLivr,
     facturebtp,
     FactCommBtp,
     PiecesRecalculs,
     UtilSoc,
     UtilTOBPiece,
     UCotraitance,
     factligneBase,
     galPatience,
     CalcOLEGenericBTP,
     BTDEMANDEDATES_TOF;

Procedure PositionneEtatAffaire(CodeAffaire, CodeEtat : String);
begin

  //FV1 - 14/03/2017 : FS#2432 - MULTIPHONE NETCOM - état du contrat = FAC quand validation définitive d'une facture
  if (Copy(CodeAffaire,0,1) = 'I') then Exit;

	ExecuteSQL('UPDATE AFFAIRE SET AFF_ETATAFFAIRE="' + CodeEtat +'",AFF_DATEMODIF="'+USDATETIME(NowH)+'" WHERE AFF_AFFAIRE="' + CodeAffaire +'"');

end;

function PositionnePieceMorteVivante (naturepiece,souche : string ;numero,indice: integer;etat:string;var EtatPrec:string):boolean;
var QQ: Tquery;
    Sql : string;
begin
  sql := 'SELECT GP_VIVANTE FROM PIECE WHERE GP_NATUREPIECEG="'+naturepiece+'" AND'
       + ' GP_SOUCHE="'+souche+'" AND '
       + ' GP_NUMERO='+inttostr(numero)+' AND GP_INDICEG='+inttostr(indice);
  QQ  := OpenSql (sql,true,-1,'',true);

  if QQ.eof then BEGIN result := false; ferme (QQ); Exit; END;

  EtatPrec := QQ.findfield('GP_VIVANTE').AsString;
  ferme (QQ);

  sql := 'UPDATE PIECE SET GP_VIVANTE="'+Etat+'" WHERE GP_NATUREPIECEG="'+naturepiece+'" AND'
       + ' GP_SOUCHE="'+souche+'" AND '
       + ' GP_NUMERO='+inttostr(numero)+' AND GP_INDICEG='+inttostr(indice);
  result :=(ExecuteSql(Sql)>0);

  sql := 'UPDATE LIGNE SET GL_VIVANTE="'+Etat+'" WHERE GL_NATUREPIECEG="'+naturepiece+'" AND'
       + ' GL_SOUCHE="'+souche+'" AND '
       + ' GL_NUMERO='+inttostr(numero)+' AND GL_INDICEG='+inttostr(indice);
  result:=(ExecuteSql(Sql)>0);

end;

Procedure AGLChangeCodeArticle ( parms: array of variant; nb: integer );
var F : TForm;
    CodeArticle,NewCode,TypeArticle,Prefixe,Libelle : string;
    i,lng : integer;
    Q : Tquery;
begin
F:=TForm(Longint(Parms[0]));
if (TFMul(F).FListe=nil) then exit;
if (PGIAskAF (TexteMessage[8], F.Caption)<>mrYes) then exit;
if TFMul(F).Fliste.AllSelected then
  BEGIN
  Q:=TFMul(F).Q;
  Q.First;
  while Not Q.EOF do
     BEGIN
     CodeArticle:=Q.FindField('GA_CODEARTICLE').AsString;
     Libelle:=Q.FindField('GA_LIBELLE').AsString;
     TypeArticle:=Q.FindField('GA_TYPEARTICLE').AsString;
     if Parms[1] = 'PREFIX' then
        begin
        if TYPEARTICLE = 'MAR' then prefixe := trim(GetParamsoc('SO_GCPREFIXEART'))
        else if (TYPEARTICLE = 'PRE') or (TYPEARTICLE = 'CTR') then prefixe := trim(GetParamsoc('SO_GCPREFIXEPRE'))
        else if (TYPEARTICLE = 'NOM') or (TYPEARTICLE = 'OUV') then prefixe := trim(GetParamsoc('SO_GCPREFIXENOM'))
        else continue; // pas de prefixe donc suivant s'il vous plait...
        // Controle prefixe Existant
        if copy(CodeArticle,1,length(prefixe))=Prefixe then continue; // pas besoin de traiter ceux qui l'ont deja
        lng := GetParamsoc('SO_GCLGNUMART');
        newCode := copy (Prefixe+CodeArticle,1,lng);
        end
     else if Parms[1] = 'CHGCODE' then
        begin
        if TYPEARTICLE = 'MAR' then prefixe := trim(GetParamsoc('SO_GCPREFIXEART'))
        else if (TYPEARTICLE = 'PRE') or (TYPEARTICLE = 'CTR') then prefixe := trim(GetParamsoc('SO_GCPREFIXEPRE'))
        else if (TYPEARTICLE = 'NOM') or (TYPEARTICLE = 'OUV') then prefixe := trim(GetParamsoc('SO_GCPREFIXENOM'))
        else prefixe := '';
        if not SaisieNewCodeArt (prefixe,TypeArticle,CodeArticle,Libelle,NewCode) then continue;
        end;
     ChangeCodificationArticle(CodeArticle,newCode);
     Q.NEXT;
     END;
  TFMul(F).Fliste.AllSelected:=False;
  END ELSE
  BEGIN
  for i:=0 to TFMul(F).Fliste.nbSelected-1 do
      begin
      TFMul(F).Fliste.GotoLeBookmark(i);
      CodeArticle:=TFMul(F).Fliste.datasource.dataset.FindField('GA_CODEARTICLE').AsString;
      Libelle:=TFMul(F).Fliste.datasource.dataset.FindField('GA_LIBELLE').AsString;
      TypeArticle:=TFMul(F).Fliste.datasource.dataset.FindField('GA_TYPEARTICLE').AsString;
      if Parms[1] = 'PREFIX' then
         begin
      // Controle prefixe Existant
         if TYPEARTICLE = 'MAR' then prefixe := trim(GetParamsoc('SO_GCPREFIXEART'))
         else if (TYPEARTICLE = 'PRE') or (TYPEARTICLE = 'CTR') then prefixe := trim(GetParamsoc('SO_GCPREFIXEPRE'))
         else if (TYPEARTICLE = 'NOM') or (TYPEARTICLE = 'OUV') then prefixe := trim(GetParamsoc('SO_GCPREFIXENOM'))
         else continue; // pas de prefixe donc suivant s'il vous plait...
         if copy(CodeArticle,1,length(prefixe))=Prefixe then continue; // pas besoin de traiter ceux qui l'ont deja
         lng := GetParamsoc('SO_GCLGNUMART');
         newCode := copy (Prefixe+CodeArticle,1,lng);
         end else if Parms[1] = 'CHGCODE' then
             begin
             if TYPEARTICLE = 'MAR' then prefixe := trim(GetParamsoc('SO_GCPREFIXEART'))
             else if (TYPEARTICLE = 'PRE') or (TYPEARTICLE = 'CTR') then prefixe := trim(GetParamsoc('SO_GCPREFIXEPRE'))
             else if (TYPEARTICLE = 'NOM') or (TYPEARTICLE = 'OUV') then prefixe := trim(GetParamsoc('SO_GCPREFIXENOM'))
             else Prefixe := '';
             if not SaisieNewCodeArt (prefixe,TypeArticle,CodeArticle,Libelle,NewCode) then continue;
             end;
      ChangeCodificationArticle(CodeArticle,newCode);
      end;
  END;
end;

function AGLIsMasterOnShared (parms : array of variant; nb : integer) : variant;
begin
  result := IsMasterOnShare(Parms[0]);
end;

Function AGLCreerPieceBTP( parms: array of variant; nb: integer ): variant;
Var CleDocAffaire : R_CLEDOC;
    Nbpiece : Integer;
    EnPa : boolean;
    StatutAffaire : String;
begin
// paramètres :
// Parms[0] : Client
// Parms[1] : Affaire
// Parms[2] : Nature de pièce
// Parms[3] : si avenant = Code affairedevis du devis de base sinon = 00
// Parms[4] : mode de traitement
// Parms[5] : Type d'affaire (A,W,I)
// Parms[6] : Saisie Cde Stock ou chantier (True or False)

	FillChar(CleDocAffaire,Sizeof(CleDocAffaire),#0) ;
	CleDocAffaire.NaturePiece:=Parms[2];
	CleDocAffaire.DatePiece:=V_PGI.DateEntree;
	enPa := false;

  StatutAffaire := Parms[5];
//
	if Pos (CledocAffaire.NaturePiece,'PBT;CBT;LBT;')  >0 then EnPa := true;
//
	if (Parms[1] = '') Or (Parms[4] = 'CRE') then
  	 NbPiece := -1
	else
  	 NbPiece := SelectPieceAffaire(Parms[1], 'AFF', CleDocAffaire);
//
  if (NbPiece > 0) then
  begin
  	 SaisiePiece(CleDocAffaire, taModif,Parms[0], Parms[1], Parms[3],false,false,false,EnPa,False, StatutAffaire,Parms[6]);
  end
  else if (Parms[3] = '') or (Parms[1] <> '') then
  begin
  	 SaisiePiece(CleDocAffaire, taCreat,Parms[0], Parms[1], Parms[3],false,false,false,EnPa,False, StatutAffaire,Parms[6]);
  end;

end;

Function DemandeAcompteOk (NaturePiece,Souche,Tiers : string ; Numero,Indice: integer; pass : boolean=false): boolean;
var XX : TAcceptationDocument;
begin
XX := TAcceptationDocument.create;
XX.naturePiece := NaturePiece;
XX.Souche := Souche;
XX.numero := numero;
XX.Indice := Indice;
XX.Tiers := Tiers;
XX.ChargeLesTobs;
//
XX.DemandeAcompte;
result := XX.result;
//
XX.free;
end;

Procedure PrepaAcceptationDevis(Codessaff, AffaireRef, naturepiece, souche:string; numero, indice:integer; Tiers:string; TOBAff :TOB;DemandeAcompte : boolean; AnnulAcceptation : boolean);
var
TOBTiers,TOBPiece:TOB;
Sql,etatprec:string;
QQ,Q1 : TQuery ;
RefPiece,REq : string;
CleDocAffaire : R_CLEDOC;
Begin
	if not AnnulAcceptation then
  begin
    if (Copy(Codessaff,16,2) <> '00') then
      begin
      // si avenant, contrôle que le devis initial est accepté
      QQ:=OpenSQL('SELECT AFF_ETATAFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="'+Copy(Codessaff,1,15)+'00"',TRUE,-1,'',true) ;
      if Not QQ.EOF then
        begin
        if (QQ.Fields[0].AsString <> 'ACP') AND (QQ.Fields[0].AsString <> 'TER') then
          begin
          PGIInfoAF ('Acceptation de l''avenant impossible car le devis initial n''est pas accepté.','Acceptation de devis');
          Ferme(QQ) ;
          exit;
          end;
        end;
     Ferme(QQ) ;
     end;
     if DemandeAcompte then
        BEGIN
        if not DemandeAcompteOk (NaturePiece,Souche,Tiers,Numero,Indice) then
           begin
           PGIInfoAF ('Acceptation impossible : Pas d''acompte défini.','Acceptation de devis');
           Exit;
           end;
        END;
     // Ajout contrôle si un devis accepté par affaire selon paramètre société
     if (GetParamSocSecur('SO_BTUNDEVISPARAFFAIRE',False)) and
        (SelectPieceAffaire(AffaireRef, 'AFF', CleDocAffaire, True) > 0) then
     begin
        PgiInfo(TraduireMemoire('Saisie impossible, un devis a déjà été accepté pour cette affaire.'));
        Exit;
     end;
  //   Modif brl 24/03 :
  //   if not PositionnePieceMorteVivante (naturepiece,souche,numero,indice,'-',etatprec) then exit;
     TOBTiers := TOB.create ('TIERS',nil,-1);
     QQ := opensql('SELECT * FROM TIERS WHERE T_TIERS="'+Tiers+'"',true,-1,'',true);
     TOBTIers.selectdb ('',QQ);
     ferme (QQ);
     if (TobTiers.GetValue('T_NATUREAUXI')='PRO') then
     begin
  //   if GetInfoParPiece(naturepiece,'GPP_PROCLI') = 'X' then
         begin
         // transformer le tiers de PRO en CLI
         TobTiers.putValue('T_NATUREAUXI','CLI');
         TOBTiers.SetAllModifie (true);
         TOBTiers.UpdateDB;
         end;
     end;
     TOBTiers.Free;

     // Lecture de l'affaire
     TOBAff.InitValeurs ;
     TOBAff.SelectDB('"'+Codessaff+'"',Nil) ;
     // Modif sur la fiche affaire du devis accepté
     TOBAff.PutValue('AFF_ETATAFFAIRE', 'ACP');
     TOBAff.PutValue('AFF_DATESIGNE', NowH);

     TOBAff.UpdateDB(false);

     //maj fiche affaire reference
     Sql := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE = "ACP" WHERE AFF_AFFAIRE='+'"'+AffaireRef+'"';
     ExecuteSQL(Sql);
  end else
  begin
   TOBPiece := TOB.Create ('PIECE',nil,-1);
   QQ := OpenSql ('SELECT * FROM PIECE WHERE GP_NATUREPIECEG="'+naturepiece+'" AND GP_SOUCHE="'+souche+'" AND '+
   								'GP_NUMERO='+InttoStr(numero)+' AND GP_INDICEG='+IntToStr(indice),true,-1,'',true);
   if not QQ.eof then
   begin
   	TOBPiece.selectDb ('',QQ);
    RefPiece := EncodeRefPiece (TOBPiece,0,false);
    RefPiece := Copy(RefPiece,1,length(RefPiece)-1); // pour enlever le ; de fin
    Req := 'SELECT DISTINCT GL_NUMERO FROM LIGNE WHERE GL_NATUREPIECEG IN ("FBT","FBP","DAC","ABT") AND GL_PIECEPRECEDENTE LIKE "'+
           RefPiece+'%"';
    Q1 := OpenSql (Req,True,-1,'',true);
    if not Q1.eof then
    begin
      TOBPiece.free;
      ferme (Q1);
      ferme(QQ);
      PgiBox (TraduireMemoire('Impossible : Une Facture existe'),'Annulation de l''acceptation de devis');
      exit;
    end else
    begin
    	ferme (Q1);
    end;
   end else
   begin
   	TOBPiece.free;
   end;
   ferme(QQ);

   TOBAff.InitValeurs ;
   TOBAff.SelectDB('"'+Codessaff+'"',Nil) ;
   if TOBAff.GetValue('AFF_PREPARE')='X' then
   begin
   	PGIInfoAF ('Impossible : Document déjà transformé','Annulation de l''acceptation de devis');
    Exit;
   end;
   // Modif sur la fiche affaire du devis accepté
   TOBAff.PutValue('AFF_ETATAFFAIRE', 'ENC');
   TOBAff.PutValue('AFF_DATESIGNE', iDate1900 );

   TOBAff.UpdateDB(false);

  end;
End;

Procedure PrepaAcceptationEtude(Codessaff, naturepiece, souche:string; numero, indice:integer; TOBAff,TobPieces :TOB);
var TobPiece:TOB;
		Sql:string;
Begin

  // Lecture de l'affaire
  TOBAff.InitValeurs ;
  TOBAff.SelectDB('"'+Codessaff+'"',Nil) ;
  // Modif sur la fiche affaire de l'étude acceptée
  if NaturePiece = 'DAP' then
  begin
		TOBAff.PutValue('AFF_ETATAFFAIRE', 'ACA')
  end else
  begin
    TOBAff.PutValue('AFF_ETATAFFAIRE', 'ACP');
  end;
  IF TOBAff.GetValue('AFF_DATESIGNE') < NowH then TOBAff.PutValue('AFF_DATESIGNE', NowH);
  TOBAff.PutValue('AFF_DATEMODIF', NowH);

  TOBAff.UpdateDB(false);

  // Remplissage de la TOB contenant les pieces selectionnées
  TobPiece := Tob.Create ('PIECE', TobPieces,-1);
  Sql := '"'+ NaturePiece
         + '";"'+ Souche
         + '";"'+ inttostr(Numero)
         + '";"'+ inttostr(Indice) + '"';

  TobPiece.selectDB (sql, Nil, False);
  if NaturePiece = 'DE' then
  begin
    // Demande de la date de livraison souhaité
  	if not DemandeDateLivraison (TobPiece) then Exit;
  end;
End;

Procedure DecisionStockReappro;
var TOBDate : TOB;
    DateFin : TdateTime;
    StWhere : String;
begin
(*
  DateFin := 0;
  TOBDate := TOB.Create ('The Date',nil,-1);
  TOBDate.AddChampSupValeur ('DATELIMITE',V_PGI.DateEntree,false);
  TRY
    TheTob := TOBDate;
    SaisirHorsInside('BTDATEPREP');
    if TheTob = nil then Exit;
    DateFin := TheTob.getValue('DATELIMITE');
  FINALLY
    TOBDate.free;
    TheTob := nil;

    stWhere := ' AND GL_DATEPIECE<="' + UsDateTime(DateFin) + '"';

    LanceEtat( 'E', 'GZF', 'BTA', True, False, False, nil, stWhere, '', False) ;
  END;
*)
    LanceEtat( 'E', 'GZF', 'BTA', True, False, False, nil, '', '', False) ;
end;

procedure AGLValideAcceptDevis( parms: array of variant; nb: integer ) ;
var F : TForm;
    Q : TQuery ;
    i :integer;
    TOBAff,TOBPieces :TOB;
    Numero, Indice : integer;
    CodeAffaire, Affaireref, NaturePiece, Souche :string;
    Tiers, Mess : string;
    DemandeAcompte,AnnulAcceptation : boolean;
    cledoc : r_cledoc;

begin
F:=TForm(Longint(Parms[0]));
DemandeAcompte := (Parms[1] = 'X');
AnnulAcceptation := (Parms[3] = 'X');
if (TFMul(F).FListe=nil) then exit;
Q:=TFMul(F).Q;
if Parms[2] = VH_GC.AFNatProposition then
begin
  Mess:=TexteMessage[10]
end else
begin
	if AnnulAcceptation then
  begin
  	Mess:=TexteMessage[11];
  end else
  begin
  	Mess:=TexteMessage[1];
  end;
end;
if (PGIAskAF (Mess, F.Caption)<>mrYes) then exit;
SourisSablier;

TOBAff:=TOB.Create('AFFAIRE',Nil,-1) ;

// on crée une TOB de toutes les études sélectionnées
TobPieces := TOB.Create ('Liste des études',NIL, -1);

try
if TFMul(F).Fliste.AllSelected then
BEGIN
  Q.First;
  while Not Q.EOF do
  BEGIN
    CodeAffaire:=Q.FindField('AFF_AFFAIRE').AsString;
    AffaireRef:=Q.FindField('GP_AFFAIRE').AsString;
    NaturePiece:=Q.FindField('GP_NATUREPIECEG').AsString;
    Souche:=Q.FindField('GP_SOUCHE').AsString;
    Numero:=Q.FindField('GP_NUMERO').AsInteger;
    Indice:=Q.FindField('GP_INDICEG').AsInteger;
    Tiers:=Q.findfield('GP_TIERS').asString;
    //
    Cledoc.Souche := Souche;
    Cledoc.NaturePiece := NaturePiece;
    Cledoc.NumeroPiece := Numero;
    Cledoc.Indice := Indice;
    //
    if not AlertDemandePrix (cledoc,TttPAcceptation) then continue;
    //
    if NaturePiece = VH_GC.AFNatAffaire then
       PrepaAcceptationDevis(CodeAffaire, AffaireRef, NaturePiece, Souche, Numero, Indice, Tiers,TOBAff,DemandeAcompte,AnnulAcceptation)
    else
       PrepaAcceptationEtude(CodeAffaire, NaturePiece, Souche, Numero, Indice, TOBAff, TobPieces);
    Q.NEXT;
  END;
  TFMul(F).Fliste.AllSelected:=False;
END
else
begin
  for i:=0 to TFMul(F).Fliste.nbSelected-1 do
  begin
    TFMul(F).Fliste.GotoLeBookmark(i);
    CodeAffaire:=TFMul(F).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
    AffaireRef:=TFMul(F).Fliste.datasource.dataset.FindField('GP_AFFAIRE').AsString;
    NaturePiece:=TFMul(F).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
    Souche:=TFMul(F).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
    Numero:=TFMul(F).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
    Indice:=TFMul(F).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
    Tiers:=Q.findfield('GP_TIERS').asString;
    //
    Cledoc.Souche := Souche;
    Cledoc.NaturePiece := NaturePiece;
    Cledoc.NumeroPiece := Numero;
    Cledoc.Indice := Indice;
    //
    if not AlertDemandePrix (cledoc,TttPAcceptation) then continue;
    //
    if NaturePiece = VH_GC.AFNatAffaire then
       PrepaAcceptationDevis(CodeAffaire, AffaireRef, NaturePiece, Souche, Numero, Indice, Tiers, TOBAff,DemandeAcompte,AnnulAcceptation)
    else
    	 PrepaAcceptationEtude(CodeAffaire, NaturePiece, Souche, Numero, Indice, TOBAff, TobPieces);

  end;
end;

// Génération des devis depuis les études
if (NaturePiece = VH_GC.AFNatProposition) then
begin
  RegroupeLesPieces(TobPieces, VH_GC.AFNatAffaire, True, False, True,0, V_PGI.DateEntree,true,false,false,true,false,'',true,'','VALIDETU');
end else if (NaturePiece = 'DE') then
begin
  RegroupeLesPieces(TobPieces, 'CC', True, False, True,0, V_PGI.DateEntree,true,false,false,true,false,'',true);
end;

finally
  TOBAff.Free; TobPieces.Free;
  TheTOB := nil;
  SourisNormale ;
end;

end;

procedure AGLClotureTechnique( parms: array of variant; nb: integer ) ;
var
F : TForm;
i,Indice :integer;
TobPieces,TobPiece,TobPieceB,TOBPiecesDAC,TOBFAC:TOB;
Numero : integer;
CodeAffaire, Affaireref, Tiers, NaturePiece, DatePiece, Souche,TypeGeneration :string;
Sql,EtatPrec,NewEtat,Req:string;
CleDocAffaire : R_CleDoc;
QQ : TQuery ;
NbMois :integer ;
DateCloture,DateFinGarantie : TDateTime;
Retour : boolean ;
LesAcomptes : TOB;
AcomptesObligatoire : boolean;
messag : string;
begin
F:=TForm(Longint(Parms[0]));
AcomptesObligatoire := false;
if (TFMul(F).FListe=nil) then exit;
if (PGIAskAF (TexteMessage[9], F.Caption)<>mrYes) then exit;
DateCloture := V_PGI.DateEntree;
if not SaisieDateCloture (DateCloture) then exit;
NbMois := ValeurI(GetParamSocSecur('SO_AFALIMGARANTI', 0));
if NbMois <> 0 then DateFinGarantie := PlusDate(DateCloture,NbMois,'M') else DateFinGarantie := PlusDate(DateCloture,12,'M');
SourisSablier;
// on crée une TOB de toutes les pieces sélectionnées
TobPieces := TOB.Create ('Liste des pieces',NIL, -1);
TobPiecesDAC := TOB.Create ('Liste des pieces acomptes',NIL, -1);
LesAcomptes := TOB.create ('Liste des acomptes',Nil,-1);
TOBFAC := TOB.Create ('LES PIECERG',nil,-1);
try
for i:=0 to TFMul(F).Fliste.nbSelected-1 do
  begin
  LesAcomptes.clearDetail;
  TFMul(F).Fliste.GotoLeBookmark(i);
  CodeAffaire:=TFMul(F).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
  AffaireRef:=TFMul(F).Fliste.datasource.dataset.FindField('GP_AFFAIRE').AsString;
  Tiers:=TFMul(F).Fliste.datasource.dataset.FindField('GP_TIERS').AsString;
  NaturePiece:=TFMul(F).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
  DatePiece:=TFMul(F).Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsString;
  Souche:=TFMul(F).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
  Numero:=TFMul(F).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
  Indice:=TFMul(F).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
  TypeGeneration:= RenvoieTypeFact(CodeAffaire); // on peut pas le recuperer dans le datasource directement

  // Saisie des avancements
  CleDocAffaire.NaturePiece:= NaturePiece;
  CleDocAffaire.DatePiece:= StrToDate(DatePiece);
  CleDocAffaire.Souche:= Souche;
  CleDocAffaire.NumeroPiece:= Numero;
  CleDocAffaire.Indice:= Indice;

  if TypeGeneration = 'DAC' then
      begin
      if PositionnePieceMorteVivante (naturepiece,souche,numero,indice,'X',EtatPrec) then
         BEGIN
         retour:=SaisieAvancementAcompte(True,AcomptesObligatoire,LesAcomptes,CleDocAffaire, taModif,Tiers, AffaireRef, '', True);
         newEtat := EtatPrec;
         if PositionnePieceMorteVivante (naturepiece,souche,numero,indice,NewEtat,EtatPrec) then
            BEGIN
            if (retour = True) then
               begin
          // Remplissage de la TOB contenant les pieces selectionnées

               TobPiece := Tob.Create ('PIECE', TobPiecesDAC,-1);
               TOBPiece.AddChampSupValeur ('AFF_OKSIZERO','-');
               TobPieceb := Tob.Create ('PIECE', TobPieces,-1);
               TOBPieceb.addChampSupValeur ('TYPEFAC',TypeGeneration,false);
(*
               Sql := '"'+ NaturePiece
                   + '";"'+ Souche
                   + '";"'+ inttostr(Numero)
                   + '";"'+ inttostr(Indice) + '"';
*)
               QQ :=OpenSql ( 'SELECT *,AFF_OKSIZERO FROM PIECE LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS ' +
                             'WHERE ' + WherePiece(CleDocAffaire, ttdPiece, False), True,-1,'',true);
               TobPiece.selectDB ('', QQ, False);
               ferme (QQ);
               (*
               if lesAcomptes.detail.count > 0 then
                  begin
                  for Indice := 0 to LesAcomptes.detail.count -1 do
                      begin
                      TOBAA := TOB.Create ('ACOMPTES',TOBPIece,-1);
                      TOBAA.dupliquer (LesAcomptes.detail[Indice],true,true);
                      end;
                  end;
               *)
               TobPieceB.dupliquer (TOBPiece,true,true);
               end;
             END;
         END;
      END ELSE
      BEGIn
      // modif brl 24/03 : en facturation directe, le devis est positionné en pièce morte lors de la clôture.
      if TypeGeneration = 'DIR' then PositionnePieceMorteVivante (naturepiece,souche,numero,indice,'-',EtatPrec);

      TobPiece := Tob.Create ('PIECE', TobPieces,-1);
      TOBPiece.AddChampSupValeur ('AFF_OKSIZERO','-');
      TOBPiece.addChampSupValeur ('TYPEFAC',TypeGeneration,false);
(*
      Sql := '"'+ NaturePiece
             + '";"'+ Souche
             + '";"'+ inttostr(Numero)
             + '";"'+ inttostr(Indice) + '"';
*)
      QQ :=OpenSql ( 'SELECT *,AFF_OKSIZERO,AFF_PREPARE FROM PIECE LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS ' +
                   'WHERE ' + WherePiece(CleDocAffaire, ttdPiece, False), True,-1,'',true);
      TOBPiece.selectDb ('',QQ,false);
      ferme (QQ);

      //Non en line
      if (TypeGeneration <> 'DIR') and
      	 (TOBPiece.getValue('GP_AFFAIRE')<> '') and
         (TOBPiece.getValue('AFF_PREPARE') <> 'X') then
         begin
	          messag := 'Le devis '+inttostr(TOBPiece.GetValue('GP_NUMERO'))+' n''est présent dans aucune prévision de chantier.'+
            					'#13#10 Confirmez-vous le traitement ?';
         	  if PgiAsk (messag) <> Mryes then
            begin
            	TOBPiece.free;
            end;
         end;
      END;
  end;

if TobPiecesDac.Detail.Count > 0 then
  begin
  // Regroupement des pieces pour génération de la facture
  if RegroupeLesPieces(TobPiecesDac, 'FBT', False, False, True,0, DateCloture,True,false,false,false,True) <> 0 then
     BEGIN
     TobPieces.Free; TOBPiecesDac.free; TOBFAC.Free;
     LesAcomptes.free;
     SourisNormale ;
     exit;
     END;
  end;

  for I := 0 to TOBPieces.detail.count -1 do
      begin
      TobPiece := TOBPieces.detail[i];
      CodeAffaire := TOBPiece.GetValue('GP_AFFAIREDEVIS');
      TypeGeneration:= RenvoieTypeFact(CodeAffaire);
      sql := 'UPDATE AFFAIRE SET AFF_ETATAFFAIRE="TER",AFF_DATEFIN="'+usdatetime(DateCloture)+
             '",AFF_DATEGARANTIE="'+UsDateTime(DateFinGarantie)+'" WHERE AFF_AFFAIRE="'+CodeAffaire+'"';
      if (ExecuteSql (Sql) > 0 ) then
         begin
         sql := 'SELECT GP_NATUREPIECEG,GP_DATEPIECE,GP_SOUCHE,GP_NUMERO,GP_INDICEG FROM PIECE WHERE GP_NATUREPIECEG="FBT" '+
                'AND GP_AFFAIREDEVIS="'+CodeAffaire+'"';
         QQ := OpenSql (sql,true,-1,'',true);
         if not QQ.eof then
            begin
            TOBFAC.ClearDetail;
            TOBFAC.LoadDetailDB ('PIECE','','',QQ,false);
            ferme (QQ);
            for Indice := 0 to TOBFAC.detail.count -1 do
                begin
                Req := 'UPDATE PIECERG SET PRG_DATEECH="'+UsDateTime(DateFinGarantie)+'"'+
                'WHERE PRG_NATUREPIECEG="'+TOBFAC.detail[Indice].GetValue('GP_NATUREPIECEG')+'" AND '+
                'PRG_DATEPIECE="'+UsDateTime(TOBFAC.detail[Indice].GetValue('GP_DATEPIECE'))+'" AND '+
                'PRG_SOUCHE="'+TOBFAC.detail[Indice].GetValue('GP_SOUCHE')+'" AND '+
                'PRG_NUMERO='+inttostr(TOBFAC.detail[Indice].GetValue('GP_NUMERO'))+' AND '+
                'PRG_INDICEG='+inttostr(TOBFAC.detail[Indice].GetValue('GP_INDICEG'));
                ExecuteSql (req)
                end;
            end else ferme (QQ);
         end;
      end;

finally
TobPieces.Free; TOBPiecesDac.free;
TOBFAC.free;
LesAcomptes.free;
SourisNormale ;
end;
end;

procedure AGLValidePrepaFactures( parms: array of variant; nb: integer ) ;
var
  F : TForm;
  i,Indice :integer;
  TobPieces,TobPiece,TOBPiecesDAC,TOBMilliemes:TOB;
  Numero : integer;
  CodeAffaire, Affaireref, Tiers, NaturePiece, DatePiece, Souche,TypeGeneration, NewNat :string;
  Sql,EtatPrec,NewEtat:string;
  CleDocAffaire : R_CleDoc;
  Retour : boolean ;
  LesAcomptes,TOBAA,TOBDATES : TOB;
  AcomptesObligatoire : boolean;
  QQ : Tquery;
  NewDate : TDateTime;
begin
F:=TForm(Longint(Parms[0]));
AcomptesObligatoire := (Parms[1]='X');
if (TFMul(F).FListe=nil) then exit;
if (PGIAskAF (TexteMessage[4], F.Caption)<>mrYes) then exit;

SourisSablier;
// on crée une TOB de toutes les pieces sélectionnées
TobPieces := TOB.Create ('Liste des pieces',NIL, -1);
TobPiecesDAC := TOB.Create ('Liste des pieces acomptes',NIL, -1);
LesAcomptes := TOB.create ('Liste des acomptes',Nil,-1);
TOBMilliemes := TOB.create ('LES MILLIEMES',Nil,-1);
TOBDATES := TOB.Create('DATE DE FAC',nil,-1);
try
for i:=0 to TFMul(F).Fliste.nbSelected-1 do
  begin
  LesAcomptes.clearDetail;
  TFMul(F).Fliste.GotoLeBookmark(i);
  CodeAffaire:=TFMul(F).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
  AffaireRef:=TFMul(F).Fliste.datasource.dataset.FindField('GP_AFFAIRE').AsString;
  Tiers:=TFMul(F).Fliste.datasource.dataset.FindField('GP_TIERS').AsString;
  NaturePiece:=TFMul(F).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
  DatePiece:=TFMul(F).Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsString;
  Souche:=TFMul(F).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
  Numero:=TFMul(F).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
  Indice:=TFMul(F).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
  TypeGeneration:=TFMul(F).Fliste.datasource.dataset.FindField('AFF_GENERAUTO').AsString;
//  TypeGeneration:= RenvoieTypeFact(CodeAffaire); // on peut pas le recuperer dans le datasource directement

  // Saisie des avancements
  CleDocAffaire.NaturePiece:= NaturePiece;
  CleDocAffaire.DatePiece:= StrToDate(DatePiece);
  CleDocAffaire.Souche:= Souche;
  CleDocAffaire.NumeroPiece:= Numero;
  CleDocAffaire.Indice:= Indice;

  if naturePiece = 'DAP' then
     begin
     TobPiece := Tob.Create ('PIECE', TobPieces,-1);
		 TOBPiece.AddChampSupValeur ('AFF_OKSIZERO','-');
 		 TOBPiece.AddChampSupValeur ('RUPTMILLIEME','');
     QQ :=OpenSql ( 'SELECT *,AFF_OKSIZERO,AFF_GENERAUTO FROM PIECE LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS ' +
                    'WHERE ' + WherePiece(CleDocAffaire, ttdPiece, False), True,-1,'',true);
     TOBPiece.selectDb ('',QQ,false);
     ferme (QQ);
     end
  // on positionne la pièce comme vivante pour saisir les avancements
  else if PositionnePieceMorteVivante (naturepiece,souche,numero,indice,'X',EtatPrec) then
     BEGIN
     retour:=SaisieAvancementAcompte(False,AcomptesObligatoire,LesAcomptes,CleDocAffaire, taModif,Tiers, AffaireRef, '', True);
     newEtat := EtatPrec;
     // on remet la pièce dans son état initial
//     if PositionnePieceMorteVivante (naturepiece,souche,numero,indice,NewEtat,EtatPrec) then
        BEGIN
        if (retour = True) then
           begin
           // Remplissage de la TOB contenant les pieces selectionnées
           if TypeGeneration = 'DAC' then
              TobPiece := Tob.Create ('PIECE', TobPiecesDAC,-1)
           else
              TobPiece := Tob.Create ('PIECE', TobPieces,-1);

      		 TOBPiece.AddChampSupValeur ('AFF_OKSIZERO','-');
      		 TOBPiece.AddChampSupValeur ('RUPTMILLIEME','');

           QQ :=OpenSql ( 'SELECT *,AFF_OKSIZERO,AFF_GENERAUTO FROM PIECE LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS ' +
                         'WHERE ' + WherePiece(CleDocAffaire, ttdPiece, False), True,-1,'',true);
           TOBPiece.selectDb ('',QQ,false);
           ferme (QQ);
//
					 TOBMilliemes.ClearDetail;
					 QQ := OPenSql ('SELECT BPM_CATEGORIETAXE,BPM_FAMILLETAXE,BPM_MILLIEME FROM BTPIECEMILIEME WHERE '+WherePiece (CledocAffaire,TTdRepartmill,false),True,-1,'',true);
           if not QQ.eof then
           begin
           	 TOBMilliemes.LoadDetailDB ('BTPIECEMILIEME','','',QQ,false);
             TOBPiece.putValue('RUPTMILLIEME',EncodeRepartTva (TOBMilliemes));
           end;
           Ferme(QQ);
//
(*           Sql := '"'+ NaturePiece
               + '";"'+ Souche
               + '";"'+ inttostr(Numero)
               + '";"'+ inttostr(Indice) + '"';
           TobPiece.selectDB (sql, Nil, False);
*)
           if lesAcomptes.detail.count > 0 then
              begin
              for Indice := 0 to LesAcomptes.detail.count -1 do
                  begin
                  TOBAA := TOB.Create ('ACOMPTES',TOBPIece,-1);
                  TOBAA.dupliquer (LesAcomptes.detail[Indice],true,true);
                  end;
              end;
           end;
         END;
     END;
  end;

// traitement des situations et factures directes
if TobPieces.Detail.Count > 0 then
  begin
    TOBDATES.AddChampSupValeur('RETOUROK','-');
    TOBDates.AddChampSupValeur('CTRLEX','X');
    TOBDates.AddChampSupValeur('TYPEDATE','Date de Facturation');
    TOBDates.AddChampSupValeur('DATFAC',iDate1900);
    TRY
      TheTOB := TOBDates;
      AGLLanceFiche('BTP','BTDEMANDEDATES','','','');
      TheTOB := nil;
    FINALLY
      if TOBDates.getValue('RETOUROK')='X' then
      begin
        NewDate := TOBDates.GetDateTime('DATFAC');
      end;
    END;
    if TOBDates.getValue('RETOUROK')<>'X' then Exit;
  // Regroupement des pieces
  if naturePiece = 'DAP' then
	begin
     if GetParamSocSecur('SO_FACTPROV',False) then NewNat := 'FPR'
     else NewNat := 'FAC';
     RegroupeLesPieces(TobPieces, NewNat, False, False, True,0, newDate)
  end else
	  RegroupeLesPieces(TobPieces, 'FBT', False, False, True,0, newDate);
  // Pour les situations, le devis initial est mort dès la première facture
  For i:=0 to TobPieces.Detail.Count-1 Do
    begin
    CodeAffaire:=Tobpieces.Detail[i].Getvalue('GP_AFFAIREDEVIS');
    TypeGeneration:= RenvoieTypeFact(CodeAffaire);
    if naturePiece = 'DAP' then
    	begin
			PositionneEtatAffaire(CodeAffaire, 'FAC');
      end;
    if TypeGeneration = 'AVA' then
      begin
      naturepiece:=Tobpieces.Detail[i].Getvalue('GP_NATUREPIECEG');
      souche:=Tobpieces.Detail[i].Getvalue('GP_SOUCHE');
      numero:=Tobpieces.Detail[i].Getvalue('GP_NUMERO');
      indice:=Tobpieces.Detail[i].Getvalue('GP_INDICEG');
//      PositionnePieceMorteVivante (naturepiece,souche,numero,indice,'-',EtatPrec);
      end;
    end;
  end;

// traitement des demandes d'acomptes
if TobPiecesDac.Detail.Count > 0 then
  begin
  // Regroupement des pieces pour demande d'acompte
  RegroupeLesPieces(TobPiecesDac, 'DAC', False, False, True,0, V_PGI.DateEntree);
  // Pour les demandes d'acomptes, le devis initial est mort dès la première facture
  For i:=0 to TobPiecesDac.Detail.Count-1 Do
    begin
    naturepiece:=TobPiecesDac.Detail[i].Getvalue('GP_NATUREPIECEG');
    souche:=TobPiecesDac.Detail[i].Getvalue('GP_SOUCHE');
    numero:=TobPiecesDac.Detail[i].Getvalue('GP_NUMERO');
    indice:=TobPiecesDac.Detail[i].Getvalue('GP_INDICEG');
//    PositionnePieceMorteVivante (naturepiece,souche,numero,indice,'-',EtatPrec);
    end;
  end;

finally
TOBDATES.free;
TobPieces.Free; TOBPiecesDac.free; TOBMilliemes.free;
LesAcomptes.free;
SourisNormale ;
end;
end;

procedure AGLAfficheAvenant( parms: array of variant; nb: integer ) ;
var
F : TForm;
i : integer;
begin
F:=TForm(Longint(Parms[0]));
TFMul(F).Fliste.EditorMode := True;

for i := 0 to TFMul(F).Fliste.Columns.Count-1 do
  begin
  if TFMul(F).Fliste.columns[i].FieldName = 'AFF_AVENANT' then
    if TFMul(F).Fliste.columns[i].Field.AsString = '00' then
    begin
        TFMul(F).Fliste.Columns[i].Field.Value := '';
    end;
  end;

end;

procedure AfterGenerationPiece (TOBOffres,TOBPieces: TOB);
var Indice,IndDetOf : integer;
    TOBTiers,TOBpiece,TOBOffre,TOBAffaireP,TOBEtude : TOB;
    Sql : string;
    QQ : TQuery;
		CleDocAffaire : R_CLEDOC;
begin
  // Mise à jour des Affaires Pieces
  for Indice := 0 to TOBOffres.detail.count -1 do
  begin
    TOBOffre := TOBOffres.detail[Indice];
    if TOBOffre.getValue('TRAITOK') = 'X' then
    begin
      for IndDetOf := 0 to TOBOffre.detail.count -1 do // niveau 1 (detail de l'appel d'offre)
      begin
        // appel de tous les documents rattachés a l'appel d'offre
        TOBEtude := TOBOffre.detail[IndDetOf];
        DecodeRefPiece (TOBEtude.GetValue('BDE_PIECEASSOCIEE'),CleDocAffaire);
        Sql := 'SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE=('+
        				 'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+WherePiece(CleDocAffaire,ttdPiece,False)+')';
        QQ := OPENSql (Sql,true,-1,'',true);
        if not QQ.eof Then
        begin
          TOBAffaireP := TOB.Create ('AFFAIRE',nil,-1);
          TOBAffaireP.selectDb ('',QQ);
          TOBAffaireP.PutValue('AFF_ETATAFFAIRE','ACP');
          TOBAffaireP.PutValue('AFF_AFFAIREREF',TOBOffre.GetValue('AFF_AFFAIRE'));
          TOBAffaireP.PutValue('AFF_AFFAIREINIT',TOBOffre.GetValue('AFF_AFFAIREINIT'));

          TOBAffaireP.UpdateDB (false);
          TOBAffaireP.Free;
        end;
        Ferme (QQ);
      end;
    end;

    TOBOffre.clearDetail;
    //
    TOBTiers := TOB.create ('TIERS',nil,-1);
    QQ := opensql('SELECT T_NATUREAUXI,T_TIERS FROM TIERS WHERE T_TIERS="'+TOBOffre.GetValue('AFF_TIERS')+'"',true,-1,'',true);
    TOBTIers.selectdb ('',QQ);
    ferme (QQ);
    if (TobTiers.GetValue('T_NATUREAUXI')='PRO') then
    begin
//      if GetInfoParPiece(VH_GC.AFNatAffaire,'GPP_PROCLI') = 'X' then
      begin
        // transformer le tiers de PRO en CLI
        TobTiers.putValue('T_NATUREAUXI','CLI');
        TOBTiers.SetAllModifie (true);
        TOBTiers.UpdateDB;
      end;
    end;
    TOBTiers.Free;
  end;
  //
end;

procedure BeforeGenerationPieces (TOBOffres,TOBPieces,TOBAffaires: TOB; Var Ok_Genere : boolean);
var Indice,IndDetOf : integer;
    TOBOffre,TOBAffaire,TOBEtude,TOBPiece,TOBAffaireP : TOB;
    Sql   : string;
    CleDocAffaire : R_CLEDOC;
    QQ    : TQuery;
    StSQL : string;
    Naff0,Naff1,Naff2,Naff3,NAvenant : string;
begin

  Ok_Genere := True;

  if TOBOffres.detail.count > 0 then
  begin
    For Indice := 0 To TOBOffres.detail.count -1 do // Niveau 0 Affaire (offres)
    begin
      TOBOffre := TOBOffres.detail[Indice];
      if TOBOffre.getValue('TRAITOK') = 'X' then
      begin
        for IndDetOf := 0 to TOBOffre.detail.count -1 do // niveau 1 (detail de l'appel d'offre)
        begin
          // appel de tous les documents rattachés a l'appel d'offre
          TOBEtude := TOBOffre.detail[IndDetOf];
          DecodeRefPiece (TOBEtude.GetValue('BDE_PIECEASSOCIEE'),CleDocAffaire);
          QQ:=OpenSQL('SELECT PIECE.*,AFF_GENERAUTO,AFF_OKSIZERO,AFF_ETATAFFAIRE AS ETATDOC FROM PIECE ' +
                      'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIRE WHERE '+WherePiece(CleDocAffaire,ttdPiece,False),True,-1,'',true) ;
          if not QQ.eof then
          begin
            TOBPiece := TOB.create ('PIECE',TOBPieces,-1);
            TOBPiece.SelectDB('',QQ) ;
            // positionne l'affaire sur lequel le document doit etre rataché
   					BTPCodeAffaireDecoupe (TOBOffre.GetString('AFF_AFFAIRE'),Naff0,Naff1,Naff2,Naff3,NAvenant,TaModif,True);
            TOBPiece.SetString('GP_AFFAIRE',TOBOffre.GetString('AFF_AFFAIRE'));
            TOBPiece.SetString('GP_AFFAIRE1',Naff1) ;
            TOBPiece.SetString('GP_AFFAIRE2',Naff2) ;
            TOBPiece.SetString('GP_AFFAIRE3',Naff3) ;
            TOBPiece.SetString('GP_AVENANT', Navenant) ;
            // ------------------------------------------------------------------
            Ferme(QQ) ;

            //FV1 : contrôle si les lignes de documents sont valorisées
            StSQL := 'select ##TOP 1##* from ligne ';
            StSQL := StSQL + 'WHERE GL_PUHT = 0 AND GL_TYPELIGNE="ART"';
            StSQL := StSQL + '  AND GL_NATUREPIECEG="' + TOBPiece.GetValue('GP_NATUREPIECEG') + '"';
            StSQL := StSQL + '  AND GL_SOUCHE="' + TOBPIece.GetValue('GP_SOUCHE') + '"';
            StSQL := StSQL + '  AND GL_NUMERO=' + inttostr(TOBPIece.GetValue('GP_NUMERO'));
            StSQL := StSQL + '  AND GL_INDICEG=' + inttostr(TOBPIece.GetValue('GP_INDICEG'));
            Ok_Genere := (Ok_Genere and not ExisteSQL(StSQL));

          end else Ferme(QQ);
        end;
      end;
    end;
    if TObAffaires.detail.count > 0 then
    begin
      // les pieces de l'affaire sont maintenant créés
      TOBAffaires.SetAllModifie (true);
      {NewOne}
      For Indice := 0 To TOBAffaires.detail.count -1 do
      begin
        if TobAffaires.detail[Indice].getValue('NEW_AFFAIRE')= 'X' then TobAffaires.detail[Indice].InsertDb(nil);
      end;
    end;
  end;
end;

procedure AGLAccepteAppelOffre( parms: array of variant; nb: integer ) ;
var
  F           : TForm;
  i,j           : integer;
  TOBOffres   : TOB;
  TOBOffre    : TOB;
  TOBAffaires : TOB;
  TOBPIeces   : TOB;
  Indice      : integer;
  CodeAffaire : string;
  QQ          : TQuery ;
  Ok_Genere   : Boolean;
begin

  F:=TForm(Longint(Parms[0]));

  if (TFMul(F).FListe=nil) then exit;
  if (PGIAskAF (TexteMessage[7], F.Caption)<>mrYes) then exit;

  SourisSablier;

  // on crée une TOB de toutes les pieces sélectionnées
  TOBOffres := TOB.Create ('Appels offre',NIL, -1);
  TOBPieces := TOB.Create ('LES PIECES',nil,-1);
  TOBAffaireS := TOB.create ('LES AFFAIRES',nil,-1);

  try
    for i:=0 to TFMul(F).Fliste.nbSelected-1 do
    begin
      TFMul(F).Fliste.GotoLeBookmark(i);
      CodeAffaire:=TFMul(F).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;

      TOBOffre := TOB.create ('AFFAIRE',TOBOffres,-1);
      QQ := opensql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="' + CodeAffaire + '"',true,-1,'',true);
      if not QQ.eof then
      begin
        TOBOffre.selectdb ('',QQ);
        TOBOffre.AddChampSupValeur ('TRAITOK','-',false);
        TOBOffre.AddChampSupValeur ('REGROUPFACTBIS',TOBOffre.GetValue('AFF_REGROUPEFACT'),false);
        TOBoffre.addchampsupvaleur('NEW_AFFAIRE', 'X');
        TOBoffre.addchampsupvaleur('AFFAIRE_INIT', 'X');
        TOBoffre.addchampsupvaleur('OLD_AFFAIRE', TOBOffre.getValue('AFF_AFFAIRE'));
        TOBoffre.addchampsupvaleur('OLD_AFFAIRE0', TOBOffre.getValue('AFF_AFFAIRE0'));
        TOBoffre.addchampsupvaleur('OLD_AFFAIRE1', TOBOffre.getValue('AFF_AFFAIRE1'));
        TOBoffre.addchampsupvaleur('OLD_AFFAIRE2', TOBOffre.getValue('AFF_AFFAIRE2'));
        TOBoffre.addchampsupvaleur('OLD_AFFAIRE3', TOBOffre.getValue('AFF_AFFAIRE3'));
        TOBoffre.addchampsupvaleur('OLD_AVENANT', TOBOffre.getValue('AFF_AVENANT'));
        //
        TOBoffre.addchampsupvaleur('NAFF_AFFAIRE', TOBOffre.getValue('AFF_AFFAIRE'));
        TOBoffre.addchampsupvaleur('NAFF_AFFAIRE0', TOBOffre.getValue('AFF_AFFAIRE0'));
        TOBoffre.addchampsupvaleur('NAFF_AFFAIRE1', TOBOffre.getValue('AFF_AFFAIRE1'));
        TOBoffre.addchampsupvaleur('NAFF_AFFAIRE2', TOBOffre.getValue('AFF_AFFAIRE2'));
        TOBoffre.addchampsupvaleur('NAFF_AFFAIRE3', TOBOffre.getValue('AFF_AFFAIRE3'));
        TOBoffre.addchampsupvaleur('NAFF_AVENANT', TOBOffre.getValue('AFF_AVENANT'));
        ferme(QQ);
        QQ := Opensql ('SELECT * FROM BDETETUDE WHERE BDE_AFFAIRE="' + CodeAffaire + '"',true,-1,'',true);
        if not QQ.eof then TOBOffre.LoadDetailDB ('BDETETUDE','','',QQ,false,true);
      end;
      ferme (QQ);
    end;

    if TobOffres.Detail.Count > 0 then
    begin
      // traitement de chaque détail
      Indice := 0;
      repeat
        TOBOffre := TOBOffres.detail[Indice];
        if GestionDetailEtude ('','',TOBOffre.GetValue('AFF_AFFAIRE'), (TOBOffre.getValue('AFF_MANDATAIRE')='X'),taModif ,TOBOffre,TatAccept) = true then
        Begin
          TOBOffre.putvalue('TRAITOK','X');
          inc (Indice);
        End else TOBOffre.free;
      until Indice>=TOBOffres.detail.count;

      if TOBOffres.detail.count = 0 then Exit;

      BeforeGenerationPieces (TOBOffres,TOBPieces, TOBAffaires,Ok_Genere);

      if not Ok_Genere then
      begin
        PGIInfo('Toutes les lignes du document ne sont pas valorisées');
      end;

      {NewOne}
      if RegroupeLesPieces(TobPieces, VH_GC.AFNatAffaire, False, False, True,0, V_PGI.DateEntree,true,false,false,false,false,'',false,'','ETUTODEV') <> 0 then
      begin
//       TOBAffaires.DeleteDB (true); // on Vire les affaires de reference
        For Indice := 0 To TOBAffaires.detail.count -1 do
        begin
          if TobAffaires.detail[Indice].getValue('NEW_AFFAIRE')= 'X' then
              TobAffaires.detail[Indice].DeleteDB (false);
        end;
      end else
      BEGIN
        AfterGenerationPiece (TOBOffres,TOBPieces);
      END;
      TheTOB := nil;
    end;
  finally
    TobOffres.Free;
    TOBPieces.free;
    TOBAFFAIRES.free;
    TheTOB := nil;
    SourisNormale ;
  end;
end;

procedure AGLPointageDemandeAct (parms: array of variant; nb: integer ) ;
var TOBAffaire,TOBTIers : TOB;
    Q : TQuery;
    Affaire : string;
begin
  TOBAffaire := TOB.Create ('AFFAIRE',nil,-1);
  TOBTiers := TOB.Create ('TIERS',nil,-1);
  Affaire := Parms[1];
  Q := Opensql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+Affaire+'"',true,-1,'',true);
  TOBAffaire.selectdb ('',Q);
  ferme (Q);
  Q := Opensql ('SELECT * FROM TIERS WHERE T_TIERS="'+TOBAffaire.GetValue('AFF_TIERS')+'"',true,-1,'',true);
  TOBTiers.selectdb ('',Q);
  ferme (Q);
  TheTOB := TOBAffaire;
  TheTob.data := TOBTiers;
  TRY
  	AGLLanceFiche('BTP','BTREGLDAC','','','ACTION=MODIFICATION') ;
  FINALLY
    TOBAffaire.free;
    TOBTiers.free;
    TheTOB := nil;
  END;
end;

procedure AjouteDebutparagFrais (TOBfrais : TOB);
var TOBL,TOBRef : TOB;
begin
  if TOBFrais.detail.count = 0 then Exit;
	TOBref := TOBFrais.detail[0];
	TOBL:=NewTobLigne(TOBFrais,1);
//	TOBL := TOB.Create ('LIGNE',TOBFrais,0);
  InitialiseLigne (TOBL,0,0);
  TOBL.putValue('GL_TYPELIGNE','DP1');
  TOBL.PutValue('GL_NATUREPIECEG', TOBRef.GetValue('GL_NATUREPIECEG'));
  TOBL.PutValue('GL_SOUCHE', TOBRef.GetValue('GL_SOUCHE'));
  TOBL.PutValue('GL_NUMERO', TOBRef.GetValue('GL_NUMERO'));
  TOBL.PutValue('GL_INDICEG', TOBRef.GetValue('GL_INDICEG'));
  TOBL.putValue('GL_TIERS',TOBref.getValue('GL_TIERS'));
//
  TOBL.PutValue('GL_TIERSFACTURE', TOBref.GetValue('GL_TIERSFACTURE'));
  TOBL.PutValue('GL_TIERSLIVRE', TOBref.GetValue('GL_TIERSLIVRE'));
  TOBL.PutValue('GL_TIERSPAYEUR', TOBref.GetValue('GL_TIERSPAYEUR'));
  TOBL.PutValue('GL_TARIFSPECIAL', TOBref.GetValue('GL_TARIFSPECIAL'));
  TOBL.PutValue('GL_TARIFTIERS', TOBref.GetValue('GL_TARIFTIERS')); //mcd 20/01/03 oubli ...
  TOBL.PutValue('GL_DATEPIECE', TOBref.GetValue('GL_DATEPIECE'));
  TOBL.PutValue('GL_NATUREPIECEG', TOBref.GetValue('GL_NATUREPIECEG'));
  TOBL.PutValue('GL_ETABLISSEMENT', TOBref.GetValue('GL_ETABLISSEMENT'));
  TOBL.PutValue('GL_FACTUREHT', TOBref.GetValue('GL_FACTUREHT'));
  TOBL.PutValue('GL_DEVISE', TOBref.GetValue('GL_DEVISE'));
  TOBL.PutValue('GL_TAUXDEV', TOBref.GetValue('GL_TAUXDEV'));
  TOBL.PutValue('GL_COTATION', TOBref.GetValue('GL_COTATION'));
  TOBL.PutValue('GL_NUMERO', TOBref.GetValue('GL_NUMERO'));
  TOBL.PutValue('GL_REGIMETAXE', TOBref.GetValue('GL_REGIMETAXE'));
  TOBL.PutValue('GL_REPRESENTANT', TOBref.GetValue('GL_REPRESENTANT'));
  TOBL.PutValue('GL_APPORTEUR', TOBref.GetValue('GL_APPORTEUR'));
  TOBL.PutValue('GL_VIVANTE', TOBref.GetValue('GL_VIVANTE'));
  TOBL.PutValue('GL_ESCOMPTE', TOBref.GetValue('GL_ESCOMPTE'));
  TOBL.PutValue('GL_REMISEPIED', TOBref.GetValue('GL_REMISEPIED'));
  TOBL.PutValue('GL_SAISIECONTRE', TOBref.GetValue('GL_SAISIECONTRE'));
  TOBL.PutValue('GL_DOMAINE', TOBref.GetValue('GL_DOMAINE'));
  {Affaires}
  TOBL.PutValue('GL_AFFAIRE', TOBref.GetValue('GL_AFFAIRE'));
  TOBL.PutValue('GL_AFFAIRE1', TOBref.GetValue('GL_AFFAIRE1'));
  TOBL.PutValue('GL_AFFAIRE2', TOBref.GetValue('GL_AFFAIRE2'));
  TOBL.PutValue('GL_AFFAIRE3', TOBref.GetValue('GL_AFFAIRE3'));
  TOBL.PutValue('GL_AVENANT', TOBref.GetValue('GL_AVENANT'));
//
  TOBL.putValue('GL_NIVEAUIMBRIC',1);
  TOBL.putValue('GL_LIBELLE',traduireMemoire('Frais de chantier'));
end;


procedure AjouteFinparagFrais (TOBfrais : TOB);
var TOBL,TOBRef : TOB;
begin
  if TOBFrais.detail.count = 0 then Exit;
	TOBref := TOBFrais.detail[0];
	TOBL:=NewTobLigne(TOBFrais,0);
//	TOBL := TOB.Create ('LIGNE',TOBFrais,-1);
  InitialiseLigne (TOBL,TOBfrais.detail.count-1,0);
  TOBL.putValue('GL_TYPELIGNE','TP1');
  TOBL.PutValue('GL_NATUREPIECEG', TOBRef.GetValue('GL_NATUREPIECEG'));
  TOBL.PutValue('GL_SOUCHE', TOBRef.GetValue('GL_SOUCHE'));
  TOBL.PutValue('GL_NUMERO', TOBRef.GetValue('GL_NUMERO'));
  TOBL.PutValue('GL_INDICEG', TOBRef.GetValue('GL_INDICEG'));
  TOBL.putValue('GL_TIERS',TOBref.getValue('GL_TIERS'));
//
  TOBL.PutValue('GL_TIERSFACTURE', TOBref.GetValue('GL_TIERSFACTURE'));
  TOBL.PutValue('GL_TIERSLIVRE', TOBref.GetValue('GL_TIERSLIVRE'));
  TOBL.PutValue('GL_TIERSPAYEUR', TOBref.GetValue('GL_TIERSPAYEUR'));
  TOBL.PutValue('GL_TARIFSPECIAL', TOBref.GetValue('GL_TARIFSPECIAL'));
  TOBL.PutValue('GL_TARIFTIERS', TOBref.GetValue('GL_TARIFTIERS')); //mcd 20/01/03 oubli ...
  TOBL.PutValue('GL_DATEPIECE', TOBref.GetValue('GL_DATEPIECE'));
  TOBL.PutValue('GL_NATUREPIECEG', TOBref.GetValue('GL_NATUREPIECEG'));
  TOBL.PutValue('GL_ETABLISSEMENT', TOBref.GetValue('GL_ETABLISSEMENT'));
  TOBL.PutValue('GL_FACTUREHT', TOBref.GetValue('GL_FACTUREHT'));
  TOBL.PutValue('GL_DEVISE', TOBref.GetValue('GL_DEVISE'));
  TOBL.PutValue('GL_TAUXDEV', TOBref.GetValue('GL_TAUXDEV'));
  TOBL.PutValue('GL_COTATION', TOBref.GetValue('GL_COTATION'));
  TOBL.PutValue('GL_NUMERO', TOBref.GetValue('GL_NUMERO'));
  TOBL.PutValue('GL_REGIMETAXE', TOBref.GetValue('GL_REGIMETAXE'));
  TOBL.PutValue('GL_REPRESENTANT', TOBref.GetValue('GL_REPRESENTANT'));
  TOBL.PutValue('GL_APPORTEUR', TOBref.GetValue('GL_APPORTEUR'));
  TOBL.PutValue('GL_VIVANTE', TOBref.GetValue('GL_VIVANTE'));
  TOBL.PutValue('GL_ESCOMPTE', TOBref.GetValue('GL_ESCOMPTE'));
  TOBL.PutValue('GL_REMISEPIED', TOBref.GetValue('GL_REMISEPIED'));
  TOBL.PutValue('GL_SAISIECONTRE', TOBref.GetValue('GL_SAISIECONTRE'));
  TOBL.PutValue('GL_DOMAINE', TOBref.GetValue('GL_DOMAINE'));
  {Affaires}
  TOBL.PutValue('GL_AFFAIRE', TOBref.GetValue('GL_AFFAIRE'));
  TOBL.PutValue('GL_AFFAIRE1', TOBref.GetValue('GL_AFFAIRE1'));
  TOBL.PutValue('GL_AFFAIRE2', TOBref.GetValue('GL_AFFAIRE2'));
  TOBL.PutValue('GL_AFFAIRE3', TOBref.GetValue('GL_AFFAIRE3'));
  TOBL.PutValue('GL_AVENANT', TOBref.GetValue('GL_AVENANT'));
//
  TOBL.putValue('GL_NIVEAUIMBRIC',1);
  TOBL.putValue('GL_LIBELLE',traduireMemoire('Total frais de chantier'));
end;

procedure IncrementeLesNiveaux  (TOBfrais : TOB) ;
var TOBL : TOB;
		Indice,TheNiveau : integer;
    TheParag : string;
begin
	for Indice := 0 to TOBfrais.detail.count -1 do
  begin
  	TOBL := TOBfrais.detail[Indice];
    if IsParagraphe (TOBL) then
    begin
       TheNiveau := TOBL.getValue('GL_NIVEAUIMBRIC');
       TheParag := Copy(TOBL.GEtValue('GL_TYPELIGNE'),1,2);
       Inc (TheNiveau);
       TOBL.PutValue('GL_NIVEAUIMBRIC',TheNiveau);
       TOBL.PutValue('GL_TYPELIGNE',TheParag+InttoStr(TheNiveau));
    end;
  end;
end;

procedure AjouteDebutFinFrais (TOBFrais : TOB);
begin
  IncrementeLesNiveaux(TOBFrais);
  AjouteDebutparagFrais (TOBfrais);
  AjouteFinparagFrais (TOBfrais);
end;

procedure AglPlannificationChantier (Parms : array of variant ; nb : integer) ;
var
TobPiece,TOBOuvrage,TOBArticles,TOBFrais,TOBOuvrageFrais,TOBBasesL,TOBSST:TOB;
TheAction : TActionFiche;
Sql:string;
CleDocAffaire,Cledoc : R_CleDoc;
//Naturepiece,souche,datepiece : string;
{Numeropiece,Indice,}IndiceOuv : integer;
Q : Tquery;
begin
  //Non en line
  if not ExistReplacePou then exit;

  IndiceOuv := 1;

  CleDocAffaire.NaturePiece:= string(Parms[1]);
  CleDocAffaire.datePiece  := strToDate(Parms[2]);
  CleDocAffaire.Souche     := string(Parms[3]);
  CleDocAffaire.NumeroPiece:= longint(Parms[4]);
  CleDocAffaire.indice     := longint(Parms[5]);

  // on crée une TOB de toutes les pieces sélectionnées
  TobPiece := TOB.Create ('PIECE',NIL, -1);
  TOBOuvrage := TOB.Create ('LIGNEOUV',nil,-1);
  TobFrais := TOB.Create ('PIECE',NIL, -1);
  TOBOuvrageFrais := TOB.Create ('LIGNEOUV',nil,-1);
  TOBBasesL := TOB.create ('LES LIGNEBASE',nil,-1);
  TOBSST := TOB.Create ('LES SOUS-TRAITS',nil,-1);
//  TOBArticles := TOB.Create ('ARTICLES',nil,-1);
  try
    // Piece
    SourisSablier;
    Q:=OpenSQL ('SELECT * FROM PIECE WHERE '+ WherePiece(CleDocAffaire,ttdPiece,False),True,-1,'',true) ;
    TobPiece.selectDB ('',Q);
    Ferme(Q) ;
    Q:=OpenSQL ('SELECT * FROM LIGNEBASE WHERE '+WherePiece(CleDocAffaire,ttdLigneBase,False),True,-1,'',true) ;
    TobBasesL.loadDetailDb ('LIGNEBASE','','',Q,false);
    Ferme(Q) ;
    LoadLesSousTraitants(CleDocAffaire ,TOBSST); 
    if PieceNonMiseAJourOptimise (TOBPiece,TOBBasesL) then BEGIN PgiBox('Veuillez recalculer la pièce préalablement.'); exit; END;
    if ControleChantierBTP (TOBPiece,BTTModif) then
    begin
      if TOBPiece.getValue('GP_AFFAIRE')<>'' then
      begin
      // Lignes de document
      Sql := MakeSelectLigneBtp (true,false,false);
      Sql := Sql + ' WHERE ' + WherePiece(CleDocAffaire, ttdLigne,false) + ' ORDER BY GL_NUMLIGNE';
      Q:=OpenSQL (SQL,True,-1,'',true) ;
      TOBPiece.LoadDetailDB ('LIGNE','','',Q,True,true) ;
      Ferme(Q) ;
      // --
      // Recup des frais detaillés
      // --
      if TOBPiece.detail.count > 0 then
      begin
//      	if TOBpiece.detail[0].GetValue('GL_PIECEORIGINE')<> '' then
        RetrouvePieceFraisBtp (Cledoc,TOBpiece,TheAction);
    		if (Cledoc.NumeroPiece <> 0) then
        begin
//          DecodeRefPiece (TOBpiece.detail[0].GetValue('GL_PIECEORIGINE'),cledoc);
//          cledoc.naturePiece := 'FRC'; // frais de chantier
          Q:=OpenSQL ('SELECT * FROM PIECE WHERE '+WherePiece(CleDoc,ttdPiece,False),True,-1,'',true) ;
          TobFrais.selectDB ('',Q);
          Ferme(Q) ;
          Sql := MakeSelectLigneBtp (true,false,false);
          Sql := Sql + ' WHERE ' + WherePiece(cledoc, ttdLigne,false) + ' ORDER BY GL_NUMLIGNE';
      		Q:=OpenSQL (SQL,True,-1,'',true) ;
      		TOBFrais.LoadDetailDB ('LIGNE','','',Q,True,true) ;
      		Ferme(Q) ;
          AjouteDebutFinFrais (TOBFrais);
      		ChargeLesOuvrages (TOBFrais,TOBOuvrageFrais,cledoc);
        end;
      end;
      // --
(*      LoadLesOuvrages (TOBPIece,TOBOuvrage,TOBArticles,cledocAffaire,IndiceOuv); *)
      ChargeLesOuvrages (TOBPIece,TOBOuvrage,cledocAffaire);
      // --
      SourisNormale ;
      SetPlannification(TOBPiece,TOBSST,TOBOuvrage,TOBFrais,TOBOuvrageFrais);
      end else
      begin
        PgiInfo ('On ne peut pas générer de prévision de chantier sur un devis non rattaché à un chantier');
      end;
    end;
  finally
    SourisNormale ;
    TobPiece.Free;
    TOBOuvrage.free;
    TobFrais.Free;
    TOBBasesL.free;
    TOBOuvrageFrais.free;
    TOBSST.Free;
//    TOBARticles.free;
  end;
end;

procedure SaisieAvancementChantier (NaturePiece, Datepiece, Souche, Tiers, AffaireRef : String; NumeroPiece, Indice : integer; Action : TActionFiche ) ;
var
  CleDocAffaire : R_CleDoc;
  EtatPrec,NewEtat:string;
begin
  CleDocAffaire.NaturePiece:= NaturePiece;
  CleDocAffaire.datePiece  := strToDate(Datepiece);
  CleDocAffaire.Souche     := Souche;
  CleDocAffaire.NumeroPiece:= NumeroPiece;
  CleDocAffaire.indice     := indice;

  // on positionne la pièce comme vivante pour saisir les avancements
  if PositionnePieceMorteVivante (naturepiece,souche,numeropiece,indice,'X',EtatPrec) then
     BEGIN
     SaisieAvancementAcompte( False,False,Nil,CleDocAffaire,Action,Tiers,AffaireRef,'',True);
     newEtat := EtatPrec;
     // on remet la pièce dans son état initial
     PositionnePieceMorteVivante (naturepiece,souche,numeropiece,indice,NewEtat,EtatPrec);
     END;

end;

procedure AglAvancementChantier (Parms : array of variant ; nb : integer) ;
begin
  SaisieAvancementChantier (Parms[1],Parms[2],Parms[3],Parms[6],Parms[7],longint(Parms[4]),longint(Parms[5]),Tamodif);
end;

{ AcceptationDocument }

constructor TAcceptationDocument.Create;
begin
  IndiceOuv := 1;
	ReinitTOBAffaires;
  CreateTobs;
  Result := false;
end;

destructor TAcceptationDocument.Destroy;
begin
  inherited;
  LibereTobs;
	ReinitTOBAffaires;
end;

procedure TAcceptationDocument.CreateTobs;
begin
  // Creation des tobs
  TobPiece := Tob.Create ('PIECE', nil,-1);
  TOBTiers := TOB.create ('TIERS',nil,-1);
  TOBAcomptes := TOB.create ('L ACOMPTE',nil,-1);
  TobPieceTrait := TOB.create('LES PIECETRAIT', nil, -1);
  TobPieceTrait_O := TOB.create('LES PIECETRAIT', nil, -1);
  TOBPieceRG := TOB.Create ('LES RETENUES',nil,-1);
  TOBBAsesRg := TOB.Create ('LES BASESRG',nil,-1);
  TOBAcomptes_O := TOB.create ('LES ACOMPTES',nil,-1);
  // ----- Piece comptabilisé
  TOBPiece_O     := TOB.Create('PIECE',nil,-1);
  TOBOuvrages    := TOB.Create('LES OUVRAGES',nil,-1);
  TOBOuvragesP 	 := TOB.Create('LES OUVRAGESP',nil,-1);
  TOBBases       := TOB.Create('LES BASES',nil,-1);
  TOBBasesL      := TOB.Create('LES BASES L',nil,-1);
  TOBEches       := TOB.Create('LES ECHES',nil,-1);
  TOBAffaireInterv := TOB.Create('LES AFFAIRE INT',nil,-1);
  TOBArticles    := TOB.Create('LES ARTICLES',nil,-1);
  TOBPorcs       := TOB.Create('LES PORCS',nil,-1);
  TOBAnaP        := TOB.Create('LES ANAP',nil,-1);
  TOBAnaS        := TOB.Create('LES ANAS',nil,-1);
  TOBSSTrait     := TOB.Create('LES SSTRAIT',nil,-1);
  TOBCPTA := CreerTOBCpta;
  // ------------------------
end;

procedure TAcceptationDocument.LibereTOBS;
begin
  FreeAndNil(TOBtiers);
  FreeAndNil(TOBPiece);
  FreeAndNil(TOBAcomptes);
  FreeAndNil(TOBAcomptes_O);
  FreeAndNil(TOBPieceRG);
  FreeAndNil(TOBBasesRG);
  FreeAndNil(TobPieceTrait);;
  FreeAndNil(TobPieceTrait_O);;
  // Comptabilisation
  TOBOuvrages.free;
  TOBOuvragesP.free;
  TOBBases.free;
  TOBBasesL.free;
  TOBEches.free;
  TOBAffaireInterv.free;
  TOBArticles.free;
  TOBPorcs.free;
  TOBAnaP.free;
  TOBAnaS.free;
  TOBSSTrait.free;
  TOBPiece_O.free;
  TOBCPTA.free;
  // ------------------
end;

procedure TAcceptationDocument.ChargeLesTObs;

	procedure  CompleteTOBArticles (TOBArticles,TOBAA : TOB);
  var II : Integer;
  		TOBA : TOB;
  begin
    ii := 0;
    repeat
      TOBA := TOBAA.detail[II];
      if TOBArticles.FindFirst(['GA_ARTICLE'],[TOBA.GetString('GA_ARTICLE')],true) = nil then
      begin
        TOBA.ChangeParent(TOBArticles,-1);
      end else
      begin
        inc(ii);
      end;
    until ii > TOBAA.detail.Count -1;
  end;

var StSelect : string;
		QQ : TQuery;
    TOBAA : TOB;
    i : Integer;
    XX : TFPatience;
begin
  XX := FenetrePatience('Gestion des Règlements',aoMilieu, False,true);
  XX.lAide.Caption := 'Chargement des données de bases...';
  XX.lcreation.visible := false ;
  XX.StartK2000 ;
  XX.Refresh;
  TOBAA := TOB.Create('LES ARTIC',nil,-1);
  TRY
    // Remplissage des tobs
    FillChar(CleDoc,Sizeof(CleDoc),#0) ;
    cledoc.NaturePiece := NaturePiece;
    cledoc.Souche := Souche;
    cledoc.NumeroPiece := Numero;
    cledoc.Indice := Indice;
    // --------------------------------
    LoadPiece (cledoc, TOBPiece);
    if TOBPiece.getValue('GP_AFFAIREDEVIS')<> '' then
    begin
      StockeCetteAffaire (string(TOBPiece.getValue('GP_AFFAIREDEVIS')));
    end;
    if TOBPiece.getValue('GP_AFFAIRE')<> '' then
    begin
      StockeCetteAffaire (string(TOBPiece.getValue('GP_AFFAIRE')));
    end;
    DEV.Code := TOBPiece.GetValue('GP_DEVISE');
    GetInfosDevise(DEV);
    DEV.Taux := TOBPiece.GetDouble('GP_TAUXDEV');
    //
    // Lecture Echéances
    QQ := OpenSQL('SELECT * FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False), True,-1, '', True);
    TOBEches.LoadDetailDB('PIEDECHE', '', '', QQ, False);
    Ferme(QQ);
    //
    fecrModifiable := EcriturePieceModifiable (TOBPiece); // Passage en compta ou pas ??
    //
    CleDoc.DatePiece:= TOBPIECE.GetValue('GP_DATEPIECE');
    // --
    QQ := opensql('SELECT * FROM TIERS WHERE T_TIERS="'+Tiers+'"',true,-1,'',true);
    TOBTIers.selectdb ('',QQ);
    ferme (QQ);

    LoadLesRetenues (TOBPiece,TOBPieceRG,TOBBasesRG);
    LoadLesAcomptes(TOBPiece,TOBAcomptes,CleDoc) ;
    TOBACOmptes_O.Dupliquer (TOBAcomptes,true,true);
    //--- chargement des pieces traitées
    LoadLesPieceTrait;
    // --------------------------------------

    // Chargement des SS traitants
    LoadLesSousTraitants(cledoc,TOBSSTRAIT);

    if fecrModifiable then
    begin
  		XX.lAide.Caption := 'Chargement des données pour mise à jour comptable...';
    	XX.Refresh;
      //
    	LoadLignes(CleDoc, TobPiece, true,false,true,false);
  		PieceAjouteSousDetail(TOBPiece,true,false,true);
      //
			TOBPiece_O.Dupliquer(TOBPiece,True,true);
      //
      // Lecture Analytiques
      LoadLesAna(TOBPiece, TOBAnaP , TOBAnaS);
      // Lecture bases Lignes
      QQ := OpenSQL('SELECT * FROM LIGNEBASE WHERE ' + WherePiece(CleDoc, ttdLigneBase, False), True,-1, '', True);
      TOBBasesL.LoadDetailDB('LIGNEBASE', '', '', QQ, False);
      Ferme(QQ);
    {$IFDEF BTP}
      OrdonnelignesBases (TOBBasesL);
    {$ENDIF}

      // Lecture bases
      QQ := OpenSQL('SELECT * FROM PIEDBASE WHERE ' + WherePiece(CleDoc, ttdPiedBase, False), True,-1, '', True);
      TOBBases.LoadDetailDB('PIEDBASE', '', '', QQ, False);
      Ferme(QQ);

      // Lecture Ports
      QQ := OpenSQL('SELECT * FROM PIEDPORT WHERE ' + WherePiece(CleDoc, ttdPorc, False), True,-1, '', True);
      TOBPorcs.LoadDetailDB('PIEDPORT', '', '', QQ, False);
      Ferme(QQ);

      LoadLesOuvrages(TOBPiece, TOBOuvrages, TOBArticles, Cledoc, IndiceOuv,nil);
      LoadLesOuvragesPlat(TOBPiece, TOBOuvragesP, Cledoc);
      //
      LoadLaTOBAffaireInterv(TOBAffaireInterv,TOBPiece.getValue('GP_AFFAIRE'));
      //
      StSelect := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                  'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                  'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE ' +
                  'WHERE GA_ARTICLE IN (SELECT DISTINCT GL_ARTICLE FROM LIGNE WHERE '+ WherePiece(cledoc,ttdLigne,False)+'  AND GL_ARTICLE <> "")'+
                  'ORDER BY GA_ARTICLE';
      QQ := OpenSQL(StSelect,True,-1,'',true);
      TOBAA.LoadDetailDB('ARTICLE','','',QQ,false);
      Ferme(QQ);
      CompleteTOBArticles (TOBArticles,TOBAA);
      //
    end;
  FINALLY
    XX.StopK2000 ;
    XX.free;
  	TOBAA.Free;
  end;
end;

Procedure TAcceptationDocument.LoadLesPieceTrait;
//var StSQL : String;
//    QQ    : TQuery;
begin
	LoadLaTOBPieceTrait (TOBPieceTrait,Cledoc,'');
  TOBPieceTrait_O.Dupliquer(TOBPieceTrait,True,true);
  (*
  //
  StSQL := 'Select * From PIECETRAIT Where ' + WherePiece(cledoc, ttdPieceTrait, True);

  QQ := OpenSQL(StSQL, True, -1, '', true);
  TobPieceTrait.LoadDetailDB('PIECETRAIT', '', '', QQ, false);
  Ferme(qq);
  //
  *)
end;



procedure TAcceptationDocument.DemandeAcompte;
var TOBACC : TOB;
    LibelleDocument : string;           
    ret : string;
begin
// Si il existe un acompte rattache inutile de demander l'acompte
if (TOBAcomptes.detail.count > 0) and (not pass) then BEGIN result:= true; Exit; END;
//
TobAcc:=Tob.Create('Les acomptes',nil,-1) ;
Tob.Create('',TobAcc,-1);
TobAcc.Detail[0].Dupliquer(TobTiers,False,TRUE,TRUE);
Tob.Create('',TobAcc.Detail[0],-1);
TobAcc.Detail[0].Detail[0].Dupliquer(TobPiece,False,TRUE,TRUE);
TheTob:=TobAcc ;
TOBAcomptes.ChangeParent(TobAcc.Detail[0].Detail[0],-1);
TheTob.data := TOBPieceRG;
LibelleDocument := 'Saisie d''acompte pour le document '+inttostr(Numero);
// ---
ret := AGLLanceFiche('BTP','BTACOMPTES','','','ACTION=CREATION;LIBELLE='+LibelleDocument) ;
TOBAcomptes.ChangeParent(nil,-1);
TobAcc.Free;
if TOBAcomptes.detail.count > 0 then
   begin
   if Transactions(ValideLaPieceAcompte,1)= oeOk then Result := true;
   end;
TheTOB := nil;
end;

function TAcceptationDocument.EcriturePieceModifiable (TOBpiece : TOB) : Boolean;
var MM : RMVT;
		Q : Tquery;
    TOBECR : TOB;
    Indice : integer;
begin
  result := false;
  TOBEcr := TOB.Create ('LES ECRITURES',nil,-1);
  TRY
    if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_TYPEPASSCPTA') = 'AUC' then Exit;
    // Contrôle période cloturée
    if (Result) and (TOBPiece.GetValue('GP_DATEPIECE')<= GetParamSocSecur('SO_DATECLOTUREPER',iDate1900)) then
    begin
      Result := False;
      Exit;
    end;
    //
    if TOBPiece.GetValue('GP_REFCOMPTABLE') = '' then Exit;
    Result := True;
    // contrôle lettrage
    MM := DecodeRefGCComptable (TOBPiece.GetValue('GP_REFCOMPTABLE'));
    Q:=OpenSQL('SELECT E_LETTRAGE,E_VALIDE,E_EXPORTE FROM ECRITURE WHERE '+WhereEcriture(tsGene,MM,False),True,-1, '', True) ;
    result := not Q.Eof;
    if result then
    begin
      TOBECR.loadDetailDb ('ECRITURE','','',Q,false);
      for Indice := 0 to TOBECR.detail.count -1 do
      begin
        if TOBECR.Detail[Indice].getValue('E_EXPORTE')='X' then
        begin
          Result := false;
          break;
        end;
      end;
    end;
    Ferme (Q);
  FINALLY
  	TOBECR.free;
  END;
end;

procedure TAcceptationDocument.GereReglements;
var TOBAcc: TOB;
  StRegle: string;
  LibelleDocument : string;
begin
  TobAcc:=Tob.Create('Les acomptes',nil,-1) ;
  TOBAcc.AddChampSupValeur('VALIDEOK','-');
  Tob.Create('',TobAcc,-1);
  TobAcc.Detail[0].Dupliquer(TobTiers,False,TRUE,TRUE);
  Tob.Create('',TobAcc.Detail[0],-1);
  TobAcc.Detail[0].Detail[0].Dupliquer(TobPiece,False,TRUE,TRUE);
  TheTob:=TobAcc;
  TheTob.data := TOBPieceRG;
  TOBpieceRG.Data := TOBPieceTrait;
  TOBPieceTrait.Data := TOBBasesRg;
  TOBBasesRg.Data := TOBSSTrait;
  TRY
    if ReglePiece then
    begin
      TOBAcomptes.ChangeParent(TobAcc.Detail[0].Detail[0],-1);
      //
      LibelleDocument := 'Saisie des réglements pour le document '+inttostr(Numero);
      // ---
      StRegle := ';ISREGLEMENT='+BoolToStr (TOBPiece.GetString('GP_NATUREPIECEG')='FBT');
      AGLLanceFiche('BTP', 'BTACOMPTES', '', '', 'ACTION=MODIFICATION;'+StRegle+';LIBELLE='+LibelleDocument);
      TOBAcomptes.ChangeParent(nil, -1);
      AcomptesVersPiece(TOBAcomptes,TobPiece);
//			GereEcheancesGC (TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TOBpieceRg,taModif,DEV,false);
    end else
    begin
      TOBSSTrait.data := TOBAcomptes;
      TOBAcomptes.data := TOBPorcs;
//      TobPieceTrait.ChangeParent(TOBAcc.Detail[0].Detail[0], -1);
      StRegle := '';
      if Fgestion = TmgCotraitance then
        StRegle := 'COTRAITANCE;'
      else
        StRegle := 'SOUSTRAITANCE;';
      AGLLanceFiche('BTP', 'BTSAIREGLTCOTRAIT', '', '', StRegle + 'ACTION=CREATION');
//      TobPieceTrait.ChangeParent(nil, -1);
    end;
  FINALLY
    TOBpieceRG.Data := nil;
    TOBPieceTrait.Data := nil;
  	TOBBasesRg.Data := nil;
    TheTOB := nil;
    if TOBAcc.GetString('VALIDEOK') = 'X' then
    begin
  		if Transactions(ValideLaPieceAcompte,0)= oeOk then Result := true;
    end else Result := false;
    TobAcc.Free;
  END;


end;

procedure TAcceptationDocument.ValideLaPieceAcompte;
var NowFutur : TDateTime;
    XX : TFPatience;
    TOBAffaire : TOB;
begin
  XX := FenetrePatience('Gestion des Règlements',aoMilieu, False,true);
  XX.lAide.Caption := 'Ecritures des données...';
  XX.lcreation.visible := false ;
  XX.StartK2000 ;
  XX.Refresh;
  TRY
    TOBAffaire := FindCetteAffaire (TOBPiece.getValue('GP_AFFAIRE'));
 		CalculeReglementsIntervenants(TOBSSTrait, TOBPiece,TOBPIECERG,TOBAcomptes,TOBPorcs,TOBPiece.getValue('GP_AFFAIRE'),TOBPieceTrait,DEV);
		GereEcheancesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPIECERG,TOBPieceTrait,TOBPorcs, TaModif, DEV, false);
    if fecrModifiable then
    begin
      NowFutur := NowH;
      if V_PGI.IoError = oeOk then DetruitCompta(TOBPiece_O, NowFutur, OldEcr, OldStk,true);
      if V_PGI.IoError = oeOk then
      begin
//        TOBeches.clearDetail;
        if not PassationComptable( TOBPiece,TOBOUvrages,TOBOuvragesP, TOBBases,TOBBasesL,
                                   TOBEches,TOBPieceTrait,TOBAffaireInterv,TOBTiers, TOBArticles,
                                   TOBCpta, TOBAcomptes, TOBPorcs, TOBPIECERG, TOBBASESRG, TOBanaP,
                                   TOBanaS,TOBSSTRAIT,nil, DEV, OldEcr,OldStk, True, false ) then
        begin
          MessageValid := 'Erreur / ecriture comptable';
          V_PGI.IoError := oeLettrage;
          Exit;
        end;
      end;
    end;
    if V_PGI.IOError = OeOk then
    begin
  // nettoyage
    if 	TOBAcomptes_O.detail.count > 0 then if not TOBAComptes_O.DeleteDB then V_PGI.IOError := OeUnknown;
    if 	TOBPieceTrait_O.detail.count > 0 then if not TOBPieceTrait_O.DeleteDB then V_PGI.IOError := OeUnknown;
    end;
    //
    if V_PGI.IOError = OeOk then ValideLesAcomptes (TOBPiece,TOBAcomptes);
    if V_PGI.IOError = OEOk then if not TOBPiece.UpdateDb (false) then V_PGI.IOError := OEUnknown;
    if (V_PGI.IOError = OEOk) and (TOBAcomptes.detail.count > 0) then if not MajMontantAcompte(TOBPiece,TOBAcomptes) then V_PGI.IOError := OEUnknown;
    if (V_PGI.IOError = OEOk) and (TOBPieceTrait.detail.count > 0) then ValideLesPieceTrait (TOBPiece,TOBAffaire,TOBPieceTrait,TOBSSTrait,DEV);
    if (V_PGI.IOError = OEOk) and (Pos(TOBPiece.GetString('GP_NATUREPIECEG'),'FBP;FBT;DAC')>0) then ReajusteSituation(TOBPiece,TOBAcomptes ,TobPieceRG,TOBBasesRg,DEV);
  FINALLY
    XX.StopK2000 ;
    XX.free;
  end;
end;

procedure AglGenereContreEtude (Parms : array of variant ; nb : integer) ;
var
TobPiece,TOBOuvrage,TOBArticles,TOBBasesL:TOB;
Sql,Req:string;
CleDocAffaire,Cledoc : R_CleDoc;
IndiceOuv : integer;
Q : Tquery;
begin
  if not ExistReplacePou then exit;
  IndiceOuv := 1;
  CleDocAffaire.NaturePiece:= string(Parms[1]);
  CleDocAffaire.datePiece  := strToDate(Parms[2]);
  CleDocAffaire.Souche     := string(Parms[3]);
  CleDocAffaire.NumeroPiece:= longint(Parms[4]);
  CleDocAffaire.indice     := longint(Parms[5]);
  // on crée une TOB de toutes les pieces sélectionnées
  TobPiece := TOB.Create ('PIECE',NIL, -1);
  TOBOuvrage := TOB.Create ('LIGNEOUV',nil,-1);
  TOBBasesL := TOB.Create('LES LIGNES BASES',nil,-1);
  try
    // Piece
    SourisSablier;
//
    Req := 'SELECT PIECE.*';
    req := req + ',AFF_GENERAUTO,AFF_OKSIZERO,AFF_ETATAFFAIRE AS ETATDOC FROM PIECE '+
                 'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS ';
    Req := Req + 'WHERE ' + WherePiece(CleDocAffaire, ttdPiece, False);
    Q := OpenSQL(Req, True,-1, '', True);
    TobPiece.SelectDb ('',Q);
//
    Ferme(Q) ;
    Q:=OpenSQL ('SELECT * FROM LIGNEBASE WHERE '+WherePiece(CleDocAffaire,ttdLigneBase,False),True,-1,'',true) ;
    TobBasesL.loadDetailDb ('LIGNEBASE','','',Q,false);
    Ferme(Q) ;
    if PieceNonMiseAJourOptimise (TOBPiece,TobbasesL) then BEGIN PgiBox('Veuillez recalculer la pièce préalablement.'); exit; END;
    if TOBPiece.getValue('GP_AFFAIRE')<>'' then
    begin
      if ControleChantierBTP (TOBPiece,BTTModif) then
      begin
        // Lignes de document
        Sql := MakeSelectLigneBtp (true,false,false) +
             ' WHERE '+WherePiece(CleDocAffaire,ttdligne,False)+
             ' ORDER BY GL_NUMLIGNE' ;
        Q:=OpenSQL (SQL,True,-1,'',true) ;
        TOBPiece.LoadDetailDB ('LIGNE','','',Q,True,true) ;
        Ferme(Q) ;
        // --
        PieceAjouteSousDetail (TOBPiece,true,false,true);
        ChargeLesOuvrages (TOBPIece,TOBOuvrage,cledocAffaire);
        // --
        SourisNormale ;
        SetContreEtude(TOBPiece,TOBOuvrage);
      end;
    end else
    begin
      PgiInfo ('On ne peut pas générer de contre-étude d''un devis non rattaché à un chantier');
    end;
  finally
    SourisNormale ;
    TobPiece.Free;
    TOBOuvrage.free;
  	TOBBasesL.free;
  end;
end;

//////////////////////////////////////////////////////////////////////////////
procedure initM3Btp();
begin
 RegisterAglFunc( 'CreerPieceBTP',False,6,AGLCreerPieceBTP);
 RegisterAglProc( 'ValideAcceptDevis',True,4,AGLValideAcceptDevis);
 RegisterAglProc( 'ValidePrepaFactures',True,1,AGLValidePrepaFactures);
 RegisterAglProc( 'AfficheAvenant',True,0,AGLAfficheAvenant);
 RegisterAglProc( 'AccepteAppelOffre',True,0,AGLAccepteAppelOffre);
 RegisterAglProc( 'ChangementCodifArticle',True,1,AGLChangeCodeArticle);
 RegisterAglProc( 'PointageDemandeAct',True,1,AGLPointageDemandeAct);
 RegisterAglProc( 'ClotureTechnique',True,0,AGLClotureTechnique);
 RegisterAglProc( 'PlannificationChantier',True,5,AGLPlannificationCHantier);
 RegisterAglProc( 'AvancementChantier',True,7,AGLAvancementChantier);
 RegisterAglProc( 'GenereContreEtude',True,5,AglGenereContreEtude);
 RegisterAglFunc( 'AGLISMasterOnShared',False,0,AGLIsMasterOnShared);

end;


procedure PleinEcran (TheForm : TForm) ;
begin
  if GetParamSocSecur('SO_SAISIEPLEINECRAN', False) then SendMessage(TheForm.Handle, WM_SYSCOMMAND, SC_MAXIMIZE, 0);
end;
Initialization
initM3Btp();
finalization
end.
