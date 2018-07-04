unit UTofBatchValid;

interface
uses  uTofAfBaseCodeAffaire,StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,Messages,HStatus,Ent1,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      Fiche, HDB, mul, DBGrids, db,dbTables,Fe_Main,
{$ENDIF}
      SaisUtil, M3VM, M3FP, HTB97, HQry, UtilGrp, FactTOB, FactTiers, FactPieceContainer ;

type
    TOF_ECmdValid_Mul = class(TOF_AFBASECODEAFFAIRE)
    private
        TWC : String;
    public
      TOBPiece, TOBAcomptes, TOBTiers, TOBArticles, TOBBases, TOBEches, TOBCpta, TOBAnaP, TOBAnaS, TOBPorcs, TOBCataLogu : TOB;
      PieceContainer: TPieceContainer;
        procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;

        procedure BatchValid;
        procedure SetValid;
//        procedure SetAllValid;

        procedure OnArgument(stArgument : String); override;
        procedure OnLoad; override;
    end;

function AcompteEnreg(TOBPiece, TOBTiers : TOB; Quiet : boolean = false) : TOB;
procedure GenereCompta(PieceContainer: TPieceContainer; DEV: RDevise);

implementation
uses Facture, FactCpta, FactGrp, FactUtil, FactComm, FactCalc, UtilCB, ETransUtil;

// procedure appellée par le bouton BVISA
procedure AGLBatchValid(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
F := TForm(Longint(Parms[0]));
if (F is TFmul) then TOTOF := TFMul(F).LaTOF
                else exit;
if (TOTOF is TOF_ECmdValid_Mul) then TOF_ECmdValid_Mul(TOTOF).BatchValid
                                else exit;
end;

Procedure TOF_ECmdValid_Mul.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff:=THEdit(GetControl('GP_AFFAIRE'))   ; Aff0:=Nil ;
Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
Tiers:=THEdit(GetControl('GP_TIERS'))   ;
END ;

procedure TOF_ECmdValid_Mul.BatchValid;
var i : integer;
begin
with TFMul(Ecran) do
   begin
   if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
      begin
      PGIBox('Veuillez sélectionner les commandes à valider', Caption);
      exit;
      end;

   if FListe.AllSelected then
      BEGIN
      if PGIAsk('Voulez-vous valider toutes les commandes ?', Caption) <> mrYes then exit;

//      if Transactions(SetAllValid, 3) <> oeOK then PGIBox('Impossible de valider toutes les commandes', Caption);
      Q.DisableControls;
      Q.First;
      InitMove(100,'');
      while not Q.EOF do
      begin
         if Transactions(SetValid, 3) <> oeOK then PGIBox('Impossible de valider la commande n° '+Q.FindField('GP_NUMERO').AsString, Caption);
         MoveCur(False);
         Q.Next;
      end;
      Q.EnableControls;

      FListe.AllSelected := false;
      END ELSE
      BEGIN
      if PGIAsk('Voulez-vous valider les commandes sélectionnées ?', Caption) <> mrYes then exit;

      InitMove(FListe.NbSelected,'');
      for i := 0 to FListe.NbSelected-1 do
         BEGIN
         FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
         Q.TQ.Seek(FListe.Row-1) ;
{$ELSE}
{$ENDIF}
         if Transactions(SetValid, 3) <> oeOK then PGIBox('Impossible de valider la commande n° '+Q.FindField('GP_NUMERO').AsString, Caption);
         MoveCur(False);
         END;
      FListe.ClearSelected;
      FiniMove;
      END;
   ChercheClick;
   end;
end;


function AcompteEnreg(TOBPiece, TOBTiers : TOB; Quiet : boolean = false) : TOB;
var X : T_GCAcompte;
    QQ : TQuery;
//    TobAc : TOB;
begin
  FillChar(X, Sizeof(X), #0);

  X.ModePaie := TOBPiece.GetValue('GP_MODEREGLE');

(*  QQ := OpenSQL('SELECT MP_JALREGLE, MP_CPTEREGLE, MP_LIBELLE FROM MODEPAIE WHERE MP_MODEPAIE="'+X.ModePaie+'"', true);
  if not QQ.EOF then begin X.JalRegle := QQ.FindField('MP_JALREGLE').AsString;
                           X.CpteRegle := QQ.FindField('MP_CPTEREGLE').AsString;
                           X.Libelle := QQ.FindField('MP_LIBELLE').AsString; end;
  Ferme(QQ); *)

  QQ := OpenSQL('SELECT MP_MODEPAIE, MP_JALREGLE, MP_CPTEREGLE, MP_LIBELLE FROM MODEREGL LEFT JOIN MODEPAIE ON MR_MP1=MP_MODEPAIE WHERE MR_MODEREGLE="'+X.ModePaie+'"', true);
  if not QQ.EOF then begin X.ModePaie := QQ.FindField('MP_MODEPAIE').AsString;
                           X.JalRegle := QQ.FindField('MP_JALREGLE').AsString;
                           X.CpteRegle := QQ.FindField('MP_CPTEREGLE').AsString;
                           X.Libelle := QQ.FindField('MP_LIBELLE').AsString; end;
  Ferme(QQ);

  if X.CpteRegle = '' then
  begin
      QQ := OpenSQL('SELECT J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL="'+X.JalRegle+'"', true);
      if not QQ.EOF then X.CpteRegle := QQ.FindField('J_CONTREPARTIE').ASString;
      Ferme(QQ);
  end;

  X.Montant := TOBPiece.GetValue('GP_TOTALTTCDEV');
  X.CBLibelle := TOBPiece.GetValue('GP_CBLIBELLE');
  X.CBInternet := TOBPiece.GetValue('GP_CBINTERNET');
  X.DateExpire := TOBPiece.GetValue('GP_CBDATEEXPIRE');
  X.TypeCarte := CodeCardType(X.CBInternet);
//  X.NumCheque :=
  X.IsReglement := true;
  result := EnregistreAcompte(TOBPiece, TOBTiers, X, Quiet);
end;

function AcompteManuel (TOBPiece, TOBTiers : TOB; ForcerRegle : Boolean) : TOB;
Var TOBAcc  : TOB ;
    StRegle : String ;
BEGIN
result := TOB.Create('', nil, -1);
TobAcc:=Tob.Create('Les acomptes',nil,-1) ; StRegle:='' ;
Tob.Create('',TobAcc,-1);
TobAcc.Detail[0].Dupliquer(TobTiers,False,TRUE,TRUE);
Tob.Create('',TobAcc.Detail[0],-1);
TobAcc.Detail[0].Detail[0].Dupliquer(TobPiece,False,TRUE,TRUE);
TheTob:=TobAcc ;
result.ChangeParent(TobAcc.Detail[0].Detail[0],-1);
if ForcerRegle then StRegle:=';ISREGLEMENT=X' ;
AGLLanceFiche('GC','GCACOMPTES','','','ACTION=MODIFICATION'+StRegle) ;
result.ChangeParent(Nil,-1);
TobAcc.Free;
END ;

{GC_NEWFACT_TP}
procedure GenereCompta(PieceContainer: TPieceContainer; DEV : RDevise);
Var OldEcr,OldStk : RMVT;
begin
FillChar(OldEcr, Sizeof(OldEcr), #0);
if not PassationComptable(PieceContainer, DEV.Decimale, OldEcr, OldStk, True)
  then V_PGI.IoError := oeLettrage;
end;

procedure TOF_ECmdValid_Mul.SetValid;
var Nature, Souche, No, Indice : String;
    PieceContainer: TPieceContainer;
    DEV : RDevise;
begin
   with TFMul(Ecran) do
   begin
     Nature := Q.FindField('GP_NATUREPIECEG').AsString;
     Souche := Q.FindField('GP_SOUCHE').AsString;
     No := Q.FindField('GP_NUMERO').AsString;
     Indice := Q.FindField('GP_INDICEG').AsString;
   end;

   TOBPiece := TOB.Create('PIECE', nil, -1);
   TOBPiece.SelectDB('"'+Nature+'";"'+Souche+'";'+No+';'+Indice, nil);
   TOBPiece.LoadDetailDB('LIGNE', '"'+Nature+'";"'+Souche+'";'+No+';'+Indice, '', nil, true);

      TOBArticles := TOB.Create('', nil, -1);
      TOBBases := TOB.Create('BASES', nil, -1);
      TOBTiers := TOB.Create('TIERS', nil, -1); TOBTiers.AddChampSup('RIB', False);
      TOBEches := TOB.Create('LES ECHEANCES', nil, -1);
      TOBCpta := TOB.Create('', nil, -1);
      TOBAnaP := TOB.Create('', nil, -1); TOBAnaS := TOB.Create('', nil, -1);
      TOBPorcs := TOB.Create('', nil, -1);
      TOBCataLogu := TOB.Create('', nil, -1);

   PieceContainer := TPieceContainer.Create;
   PieceContainer.InitTobs(Self);
   PieceAjouteSousDetail(TOBPiece);
   DEV.Code := TOBPiece.GetValue('GP_DEVISE'); GetInfosDevise(DEV);
   if RemplirTOBTiers(TOBTiers, TOBPiece.GetValue('GP_TIERS'), Nature, False) <> trtOk then V_PGI.IoError := oeUnknown;
   UG_AjouteLesArticles(TOBPiece,TOBArticles,TOBCpta,TOBTiers,TOBAnaP,TOBAnaS,TOBCataLogu,False) ;

   CalculFacture(PieceContainer);
   CalculeSousTotauxPiece(TOBPiece);

   if IsModeReglChq(TOBPiece.GetValue('GP_MODEREGLE'))
     then TOBAcomptes := AcompteManuel(TOBPiece, TOBTiers, true) // Doit on passer true ? (-> ISREGLEMENT=X)
     else TOBAcomptes := AcompteEnreg(TOBPiece, TOBTiers);

   if TOBAcomptes = nil then TOBAcomptes := TOB.Create('', nil, -1)
                        else AcomptesVersPiece(PieceContainer);

   {Echéances}
   TOBEches.ClearDetail;
   GereEcheancesGC(PieceContainer, False);

   TOBPiece.SetAllModifie(True); TOBPiece.SetDateModif(NowH);
   TOBBases.SetAllModifie(True);
   TOBEches.SetAllModifie(True);
   TOBAcomptes.SetAllModifie(True);
   TOBTiers.SetAllModifie(True);

   if V_PGI.IoError=oeOk then GenereCompta(PieceContainer, DEV);
   if V_PGI.IoError=oeOk then TOBBases.InsertOrUpdateDB;
   if V_PGI.IoError=oeOk then TOBEches.InsertOrUpdateDB;
   if V_PGI.IoError=oeOk then TOBAcomptes.InsertOrUpdateDB;
   if V_PGI.IoError=oeOk then TOBPiece.InsertOrUpdateDB;

      TOBArticles.Free;
      TOBBases.Free;
      TOBTiers.Free;
      TOBEches.Free;
      TOBAcomptes.Free;
      TOBCpta.Free;
      TOBAnaP.Free;
      TOBAnaS.Free;
      TOBPorcs.Free;
      TOBCataLogu.Free;

   TOBPiece.Free;
   PieceContainer.Free;
end;

{procedure TOF_ECmdValid_Mul.SetAllValid;
begin

end;}

procedure TOF_ECmdValid_Mul.OnArgument(stArgument : String);
Var CC : THValComboBox ;
begin
inherited;
if stArgument <> '' then
   begin
   SetControlText('GP_VENTEACHAT',stArgument);
   THValComboBox(Ecran.FindComponent('GP_NATUREPIECEG')).Plus:='GPP_VENTEACHAT="'+stArgument+'"';
   end;
if Not(ctxGCAFF in V_PGI.PGIContexte) then  SetControlVisible ('TGP_AFFAIRE1', False);
CC:=THValComboBox(GetControl('GP_DOMAINE')) ; if CC<>Nil then PositionneDomaineUser(CC) ;
CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
end;

procedure TOF_ECmdValid_Mul.OnLoad;
begin
inherited;
TWC := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));
end;


initialization
RegisterClasses([TOF_ECmdValid_Mul]);
RegisterAGLProc('BatchValid', true, 0, AGLBatchValid);
end.
