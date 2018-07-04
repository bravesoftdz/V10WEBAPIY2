{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 04/09/2002
Modifié le ... : 04/09/2002
Description .. : Epuration des documents morts
Mots clefs ... :
*****************************************************************}
unit PieceEpure_Tof;

{========================== JLD 15 mai 20032 ====================}
(* Attention, il faut tout reprendre !!!
   - Par rapport à l'incidence des natures de pièce
   - Par rapport au désimpact ou non de ces épurations
   - Par rapport à la cohérence vis à vis des autres produits de l'intégré
   - Par rapport aux systèmes de clôture (anciens ou nouveaux)
   - Par rapport à l'emplacement de ces fonctions deans les menus
   - Par rapport au code et script (Appels et manipulations de contrôles non existants, nqtures de pièce en dur, ...)
*)

interface
uses  uTofAfBaseCodeAffaire,StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,Messages,HStatus,Ent1, UTofTTPeriode,
      UtilMulTrt,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      Fiche, HDB, mul, DBGrids, db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
      M3VM, M3FP, HTB97, HQry, EntGC, UtilPGI, AssistPieceEpure, FactComm,
      FactUtil, FactCpta, SaisUtil, SaisComm, FactTOB,uEntCommun,UtilTOBPiece;

function GCLanceFiche_PieceEpure (Nat,Cod : String ; Range,Lequel,Argument : string) : string;

type
    TOF_GCPieceEpure_Mul = class(TOF_AFBASECODEAFFAIRE)
    private
        BOuvrir : TToolBarButton97;
        NaturePieceg : THValComboBox;
    public
        TobPiece : TOB;
        procedure OnArgument(stArgument : String); override;
        procedure BOuvrirClick (Sender : TObject);
        procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
        function  EpureLaPiece(TobP : TOB) : boolean;
        function  ChargeLesPieces (var iListeNumeros : array of integer) : boolean;
        function  ExisteLignePrecedente ( CD : R_CleDoc ) : Boolean ;
        function  GenereInfoLigne(TOBP : TOB ; CleDoc : array of R_CleDoc ; CD : R_CleDoc) : boolean;
        procedure NaturePieceChange (Sender : TObject);
    end;

Const icoY = '#ICO#37'; icoN = '#ICO#35';
      icoInfo = '#ICO#7';

      TexteMessage: array[1..5] of string 	= (
          {1}   'Les pièces ne peuvent être purgées au delà de la date de dernière clôture de stock.'
          {2}  ,'Vous devez solder les quantités en reliquat présentes sur cette pièce.'
          {3}  ,'Cette pièce n''est pas épurable, elle est liée à une écriture comptable.'
          {4}  ,'Cette pièce possède une ou plusieurs pièces précédentes. Vous devez au préalable les supprimer.'
          {5}  ,''
              );

      MaxPiece = 2000; //nbre de pièces "sélectionnables"

implementation

function GCLanceFiche_PieceEpure (Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result := AGLLanceFiche (Nat,Cod,Range,Lequel,Argument);
end;

Procedure TOF_GCPieceEpure_Mul.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
begin
Aff:=THEdit(GetControl('GP_AFFAIRE'))   ;
Aff0:=Nil ;
Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
Tiers:=THEdit(GetControl('GP_TIERS'))   ;
end ;

function TOF_GCPieceEpure_Mul.ChargeLesPieces (var iListeNumeros : array of integer) : boolean;
var iInd : integer;
    TOBP : TOB;
    iIndNumero, iIndBis : integer;
begin
Result := True;
for iIndNumero := 0 to 999 do
    begin
    iListeNumeros [iIndNumero] := 0;
    end;
iIndNumero := 0;

TraiteEnregMulTable (TFMul(Ecran),'SELECT GP_NATUREPIECEG, GP_DATEPIECE, GP_SOUCHE, ' +
                      'GP_NUMERO, GP_INDICEG, GP_TIERS, GP_REFCOMPTABLE FROM PIECE ','GP_NATUREPIECEG;GP_NUMERO;GP_SOUCHE;GP_INDICEG','PIECE','GP_NUMERO','PIECE',TobPiece,True,MaxPiece);
InitMove(TobPiece.Detail.Count, '');
for iInd := 0 to TobPiece.Detail.Count - 1 do
begin
  MoveCur (False);
  TOBP := TobPiece.Detail[iInd];
  TOBP.AddChampSupValeur('EPURABLE',False);
  TOBP.AddChampSupValeur('ICOEPURE',IcoN);
  TOBP.AddChampSupValeur('ICOINFO',IcoInfo);
  TOBP.AddChampSupValeur('BLOCNOTE','');
  TOBP := TobPiece.Detail [iInd];
  iIndBis := 0;
  while (iIndBis < 1000) and (iListeNumeros [iIndBis] <> 0) and
        (iListeNumeros [iIndBis] <> TOBP.GetValue('GP_NUMERO')) do
      begin // pour ne pas traiter deux fois une piece vu qu'on ne s'occupe pas des indiceg
      iIndBis := iIndBis + 1;
      end;
  if EpureLaPiece (TOBP) then
      begin
      if (iIndBis = 1000) or (iListeNumeros [iIndBis] = 0) then
          begin
          iListeNumeros [iIndNumero] := TOBP.GetValue('GP_NUMERO');
          iIndNumero := iIndNumero + 1;
          TOBP.PutValue('EPURABLE',True);
          TOBP.PutValue('ICOEPURE',IcoY); TOBP.PutValue('ICOINFO','');
          end;
      end;
end;
FiniMove;
end;

procedure TOF_GCPieceEpure_Mul.NaturePieceChange (Sender : TObject);
begin
  if (TFMul(Ecran).FListe.NbSelected > 0) or (TFMul(Ecran).FListe.AllSelected) then
  begin
    TFMul(Ecran).ChercheClick;
  end;
end;

procedure TOF_GCPieceEpure_Mul.BOuvrirClick (Sender : TObject);
var iNumEvt, iIndNumero : integer;
    TSqlJournal : TQuery;
    FAssistPE : TFAssistantPieceEpure;
    iListeNumeros : array [1..1000] of integer;
    stListeNumeros : string;
begin
with TFMul(Ecran) do
    begin
    if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
        begin
        PGIBox('Veuillez sélectionner les pièces à épurer', Caption);
        end else
        begin
        TobPiece := Tob.Create ('PIECE', Nil, -1);
        if TobPiece <> nil then
            begin
            if ChargeLesPieces (iListeNumeros) then
                begin
                FAssistPE := TFAssistantPieceEpure.Create (Application);
                with FAssistPE do
                    begin
                    TobJournal := TOB.Create('JNALEVENT', Nil, -1) ;

                    TobJournal.PutValue ('GEV_TYPEEVENT', 'EPU');
                    TobJournal.PutValue ('GEV_LIBELLE',
                                         GetInfoParPiece (GetControlText ('GP_NATUREPIECEG'),
                                                          'GPP_LIBELLE'));
                    TobJournal.PutValue ('GEV_DATEEVENT', Date);
                    TobJournal.PutValue ('GEV_UTILISATEUR', V_PGI.User);

                    TLBRecapitulatif.Items.Clear;
                    TLBRecapitulatif.Items.Add ('Epuration des "' +
                                                    GetInfoParPiece (GetControlText ('GP_NATUREPIECEG'),
                                                                                     'GPP_LIBELLE') + '"');
                    TLBRecapitulatif.Items.Add ('');

                    TLBRecapitulatif.Items.Add ('Critères de sélection : ');
                    TLBRecapitulatif.Items.Add ('   Date comprise entre le ' +
                                                    GetControlText ('GP_DATEPIECE') +
                                                    ' et le ' + GetControlText ('GP_DATEPIECE_'));
                    if GetControlText ('GP_TIERS') <> '' then
                        begin
                        TLBRecapitulatif.Items.Add ('   Tiers : ' + GetControlText ('GP_TIERS') + ' - ' +
                                                        RechDom ('GCTIERS', GetControlText ('GP_TIERS'), False));
                        end;
                    if (GetControlText ('GP_NUMERO') <> '') or (GetControlText ('GP_NUMERO_') <> '') then
                        begin
                        TLBRecapitulatif.Items.Add ('   du numéro ' + GetControlText ('GP_NUMERO') + ' au ' +
                                                        GetControlText ('GP_NUMERO_'));
                        end;
                    TLBRecapitulatif.Items.Add ('');
                    TLBRecapitulatif.Items.Add ('Numéro des "' +
                                                    GetInfoParPiece (GetControlText ('GP_NATUREPIECEG'),
                                                                                     'GPP_LIBELLE') +
                                                    '" épurés(ées)');
                    iIndNumero := 1;
                    stListeNumeros := '';
                    while iListeNumeros [iIndNumero] <> 0 do
                        begin
                        if stListeNumeros <> '' then
                            stListeNumeros := stListeNumeros + ', ';
                        stListeNumeros := stListeNumeros + IntToStr (iListeNumeros [iIndNumero]);
                        iIndNumero := iIndNumero + 1;
                        if Length (stListeNumeros) > 50 then
                            begin
                            TLBRecapitulatif.Items.Add (stListeNumeros);
                            stListeNumeros := '';
                            end;
                        end;
                    if stListeNumeros <> '' then
                        TLBRecapitulatif.Items.Add (stListeNumeros);
                    TLBRecapitulatif.Items.Add ('');

                    TobPiece.PutGridDetail(GListePieces, False,
                                            True, 'ICOEPURE;GP_NUMERO;GP_DATEPIECE;GP_TIERS;ICOINFO',True);
                    GListePieces.Cells [0, 0] := '';
                    GListePieces.Cells [1, 0] := ChampToLibelle ('GP_NUMERO');
                    GListePieces.Cells [2, 0] := ChampToLibelle ('GP_DATEPIECE');
                    GListePieces.Cells [3, 0] := ChampToLibelle ('GP_TIERS');
                    GListePieces.Cells [4, 0] := TraduireMemoire('Info');

                    TobPieceAssist := TobPiece;
                    stNaturePieceG := GetControlText ('GP_NATUREPIECEG');
                    ShowModal;

                    TOBJournal.PutValue('GEV_BLOCNOTE', TLBRecapitulatif.Items.Text);

                    iNumEvt := 0;
                    TSqlJournal := OpenSQL ('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1,'',true);
                    if Not TSqlJournal.EOF then iNumEvt := TSqlJournal.Fields[0].AsInteger;
                    Ferme (TSqlJournal);
                    Inc (iNumEvt);
                    TOBJournal.PutValue ('GEV_NUMEVENT', iNumEvt);
                    TOBJournal.InsertDB (Nil) ;
                    TobJournal.Free;

                    GListePieces.VidePile (False);
                    TLBRecapitulatif.Items.Clear;
                    FAssistPE.Free;
                    end;
                end;
            TobPiece.Free;
            end;

        if FListe.AllSelected then FListe.AllSelected:=False
        else FListe.ClearSelected;
        bSelectAll.Down := False ;

        ChercheClick;
        end;
    end;
end;

function TOF_GCPieceEpure_Mul.EpureLaPiece(TOBP : TOB) : boolean;
var TSql : TQuery;
    iInd : integer;
    NaturePiece : string;
    CleDoc : array of R_CleDoc;
    CD : R_CleDoc;
    TOBL : TOB;
    MM : RMVT ;
begin
   Result := True;  NaturePiece := TOBP.GetValue('GP_NATUREPIECEG');
   // la piece affecte t elle le physique
   if (Pos ('PHY', GetInfoParPiece (NaturePiece, 'GPP_QTEMOINS')) > 0) or
      (Pos ('PHY', GetInfoParPiece (NaturePiece, 'GPP_QTEPLUS')) > 0) then
   begin
       // si oui, sa date < dernière date de cloture
       if StrToDateTime (GetControlText ('DATEMAX')) < TOBP.GetValue('GP_DATEPIECE') then
       begin Result := False ; TOBP.PutValue('BLOCNOTE',TraduireMemoire(TexteMessage[1])); exit; end;
   end;

   // cette pièce est elle vivante pour n'importe quel indice ?
   TSql := OpenSql ('SELECT COUNT(*) FROM PIECE WHERE GP_NATUREPIECEG="' +
                           NaturePiece + '" AND GP_SOUCHE="' +
                           TOBP.GetValue('GP_SOUCHE') + '" AND GP_NUMERO=' +
                           IntToStr(TOBP.GetValue('GP_NUMERO')) + ' AND GP_VIVANTE="X"', True,-1,'',true);
   iInd := TSql.Fields [0].AsInteger;
   Ferme (TSql);
   if iInd > 0 then begin Result := False ; TOBP.PutValue('BLOCNOTE',TraduireMemoire(TexteMessage[2])); exit; end;

    //écritures comptables non purgées
   if TOBP.GetValue('GP_REFCOMPTABLE') <> '' then
   begin
      MM:=DecodeRefGCComptable(TOBP.GetValue('GP_REFCOMPTABLE')) ;
      TSql:=OpenSQL('SELECT COUNT(*) FROM ECRITURE WHERE '+WhereEcriture(tsGene,MM,False),True,-1,'',true) ;
      if not TSql.EOF then iInd := TSql.Fields [0].AsInteger;
      Ferme (TSql);
      if iInd > 0 then begin Result := False ; TOBP.PutValue('BLOCNOTE',TraduireMemoire(TexteMessage[3])); exit; end;
   end;
    //pièces précédentes existentes    -    en dernier test on charge les lignes
    TSql:=OpenSQL('SELECT GL_INDICEG,GL_NATUREPIECEG,GL_NUMERO,GL_NUMLIGNE,GL_PIECEPRECEDENTE,GL_SOUCHE FROM LIGNE '
        + 'WHERE GL_NATUREPIECEG="' + TOBP.GetValue('GP_NATUREPIECEG') + '"'
        + 'AND GL_SOUCHE="' + TOBP.GetValue('GP_SOUCHE') + '" '
        + ' AND GL_NUMERO='+ IntToStr(TOBP.GetValue('GP_NUMERO'))
        + ' AND GL_INDICEG='+ IntToStr(TOBP.GetValue('GP_INDICEG')) + ' ORDER BY GL_NUMLIGNE',True,-1,'',true) ;

    TOBP.LoadDetailDB('LIGNE','','',TSql,False,False) ;
    Ferme (TSql);
    setlength(CleDoc,0);
    for iInd := 0 to TOBP.Detail.Count -1 do
    begin
       TOBL := TOBP.Detail[iInd];
       if TOBL.GetValue('GL_PIECEPRECEDENTE') = '' then continue;
       DecodeRefPiece (TOBL.GetValue('GL_PIECEPRECEDENTE'),CD);
       if ExisteLignePrecedente(CD) then
          if GenereInfoLigne(TOBP,CleDoc,CD) then
          begin
          setlength(CleDoc,Length(CleDoc)+1);
          CleDoc[High(CleDoc)] := CD;
          end;
    end;
    TOBP.ClearDetail;    CleDoc := nil;
    if TOBP.GetValue('BLOCNOTE') <> '' then Result := False;
end;

procedure TOF_GCPieceEpure_Mul.OnArgument(stArgument : String);
Var CC : THValComboBox ;
    StPlus, stVenteAchat : string;
begin
inherited;
BOuvrir := TToolBarButton97(GetControl ('BOuvrir'));
BOuvrir.OnClick := BOuvrirClick;

NaturePieceG := THValComboBox (GetControl ('GP_NATUREPIECEG'));
NaturePieceg.OnChange := NaturePieceChange;

stVenteAchat := stArgument; // appel avec Argument égal VEN ou ACH
StPlus := 'AND (GPP_VENTEACHAT="' + stVenteAchat + '")';
if stVenteAchat = 'VEN' then
    begin
    SetControlCaption ('TGP_TIERS', 'Client');
    end else
    begin
    if Not (ctxScot in V_PGI.PGIcontexte) then SetControlText ('GP_NATUREPIECEG', 'CF');
    setControlproperty ('GP_TIERS', 'DataType', 'GCTIERSFOURN');
    setControlproperty ('GP_TIERS_', 'DataType', 'GCTIERSFOURN');
    SetControlText ('TGP_TIERS', 'Fournisseur');
    end;

SetControlProperty ('GP_NATUREPIECEG', 'Plus', StPlus);
if stVenteAchat = 'VEN' then SetControlText ('GP_NATUREPIECEG', 'DE')
else SetControlText ('GP_NATUREPIECEG', 'CF');
{$IFNDEF BTP}
if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
  SetControlVisible ('TGP_AFFAIRE1', False);
{$ENDIF}
CC := THValComboBox (GetControl ('GP_DOMAINE'));
if CC <> Nil then PositionneDomaineUser (CC);
CC := THValComboBox (GetControl ('GP_ETABLISSEMENT'));
if CC <> Nil then PositionneEtabUser (CC) ;
end;

procedure AGLRecupDateMaximum (Parms : Array of Variant; Nb : Integer);
var stNaturePiece, stDateMax : string;
    F : TForm;
    THDate : THEdit;
    EstAvoir : boolean;
begin
F := TForm(Longint(Parms[0]));
stNaturePiece := String(Parms[1]);
EstAvoir := (GetInfoParPiece(stNaturePiece, 'GPP_ESTAVOIR') = 'X');
if (Pos ('PHY', GetInfoParPiece (stNaturePiece, 'GPP_QTEMOINS')) > 0) or
   (Pos ('PHY', GetInfoParPiece (stNaturePiece, 'GPP_QTEPLUS')) > 0)  or EstAvoir then
    begin
    if (stNaturePiece = 'FAC') or (EstAvoir) then
        begin
        stDateMax := DateToStr (FinDeMois(PlusMois(VH^.Encours.Deb,-1))); //RechDateFinExo - 1);
        end else
        begin
        stDateMax := DateToStr (RechDerniereCloture);
        end;
    end else
    begin
    stDateMax := DateToStr (V_PGI.DateEntree);
    end;
THDate := THEdit(F.FindComponent ('DATEMAX'));
THDate.Text := stDateMax;
THEdit(F.FindComponent ('GP_DATEPIECE')).Text := stDateMax;
THEdit(F.FindComponent ('GP_DATEPIECE_')).Text := stDateMax;
end;

function TOF_GCPieceEpure_Mul.ExisteLignePrecedente ( CD : R_CleDoc ) : Boolean ;
var QQ : TQuery ;
begin
Result := False;
QQ:=OpenSQL('SELECT COUNT(*) FROM PIECE WHERE '+WherePiece(CD,ttdPiece,True),True,-1,'',true) ;
if (not QQ.EOF) and (QQ.Fields [0].AsInteger > 0) then Result := true;
Ferme(QQ) ;
end;

function TOF_GCPieceEpure_Mul.GenereInfoLigne(TOBP : TOB ; CleDoc : array of R_CleDoc ; CD : R_CleDoc) : boolean;
var Msg : string;
    iInd : integer;
begin
  Result := False;
  if TOBP = nil then exit;
  for iInd := 0 to High(CleDoc) do
  begin
    if (CleDoc[iInd].NaturePiece = CD.NaturePiece) and
       (CleDoc[iInd].Souche = CD.Souche) and
       (CleDoc[iInd].NumeroPiece = CD.NumeroPiece) and
       (CleDoc[iInd].Indice = CD.Indice) then exit;
  end;
  Msg := TOBP.GetValue('BLOCNOTE');
  if Msg = '' then Msg := TraduireMemoire(TexteMessage[4]) + ' : ';
  Msg := Msg + '| - ' + RechDom('GCNATUREPIECEG',CD.NaturePiece,False)
      + ' N° ' + intToStr(CD.NumeroPiece) + ' Souche '
      + CD.Souche + ' Indice ' + intToStr(CD.Indice);
  TOBP.PutValue('BLOCNOTE',Msg);
  Result := True
end;

initialization
RegisterClasses ([TOF_GCPieceEpure_Mul]);
RegisterAGLProc ('RecupDateMaximum', True, 2, AGLRecupDateMaximum);
end.
