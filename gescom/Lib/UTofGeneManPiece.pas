unit UTofGeneManPiece;

interface
uses {$IFDEF VER150} variants,{$ENDIF}  uTofAfBaseCodeAffaire,StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, UTOM,AGLInit,
      Utob,HDB,Messages,HStatus,Ent1,Hqry,HTB97,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      Fiche,mul, DBGrids,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
{$IFDEF BTP}
      BTPUtil,
{$ENDIF}
      M3FP,SaisUtil,FactGrpMan,FactUtil,FactComm,Facture, EntGC,uEntCommun,UtilTOBPiece;

Type

     TOF_GCGeneManPiece = Class (TOF_AFBASECODEAFFAIRE)
     private
        GLISTE : THGRID ;
        PieceGenere, PieceRegroupe,ColName : string ;
        TOBSelect : TOB;
        BSELECTALL : TToolbarButton97;
        TypeTraitement : string;
        procedure InitGrille ;
        procedure GenererPiece ;
        procedure VoirPiece ;
        Procedure AffichageTotal ;
        procedure BSelectAllClick(Sender: TObject);
        Procedure ChargeChampsRupt (NatPiece: string; var ChampsRupt : TStrings);
        function  AjustChampToStr(Champ : string; V1 : Variant) : string;
        function  FabriqueWherePiece (Tob_Piece : TOB) : string;
        procedure ChargeLesPiecesRegroupables ;
        procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;
    function DemandeDatesLivraison(var DateLiv: TDateTime): boolean;
     public
     		procedure FlisteDblClick (Sender : Tobject);
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnLoad; override;
        procedure OnUpdate ; override;
        procedure OnClose; override;
        procedure FlipSelection(Sender: TObject);
     END ;

procedure AGLGenererManPiece(parms:array of variant; nb: integer ) ;
procedure AGLGeneManPiece_VoirPiece(parms:array of variant; nb: integer ) ;

implementation
uses facttob;

Procedure TOF_GCGeneManPiece.NomsChampsAffaire ( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ;
BEGIN
Aff:=THEdit(GetControl('GP_AFFAIRE'))   ; Aff0:=Nil ;
Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
Tiers:=THEdit(GetControl('GP_TIERS'))   ;
END ;

Procedure TOF_GCGeneManPiece.OnArgument (Arguments : String ) ;
Var CC : THValComboBox ;
    St,Arg : string ;
{$IFDEF BTP}
    NaturePiece, CC2 : THMultiValComboBox;
    XX_WHERE : THedit;
{$ENDIF}
begin
inherited ;
Arg:='';
while Arguments<>'' do
      BEGIN
      St:=Uppercase(ReadTokenSt(Arguments)) ;
      if St='VENTE' then SetControlProperty('GP_TIERS','DataType','GCTIERSCLI')
        else if St='ACHAT' then SetControlProperty('GP_TIERS','DataType','GCTIERSFOURN')
{$IFDEF BTP}
        else if (St='LIVRAISON') or (St='FACTURATION') or (st = 'RETOURFOU') or (st = 'LIVRAISONCLI') then TypeTraitement := st
{$ENDIF}
          else Arg:=Arg+St+';';
      END;
//PieceRegroupe:=Uppercase(ReadTokenSt(Arguments)) ;
//PieceGenere:=Uppercase(ReadTokenSt(Arguments)) ;
PieceRegroupe:=Uppercase(ReadTokenSt(Arg)) ;
PieceGenere:=Uppercase(ReadTokenSt(Arg)) ;
{$IFDEF BTP}
if ecran.name = 'BTGROUPEMANPIECE' then
begin
	 THdbGrid(GetCOntrol('FLISTE')).ondblclick := FlisteDblClick;
   Naturepiece := THMultiValComboBox (GetControl('GP_NATUREPIECEG'));
   XX_WHERE := THEdit(GetControl('XX_WHERE'));
   if (TypeTraitement = 'LIVRAISON') then
   begin
      TCheckBox(GetControl('GP_VIVANTE')).Checked :=true;
      ecran.caption := TraduireMemoire ('Génération des réceptions fournisseurs manuelles');
      Naturepiece.plus := 'AND ((GPP_NATUREPIECEG="CF") OR (GPP_NATUREPIECEG="CFR"))';
      XX_WHERE.Text  :='((GP_NATUREPIECEG="CF") OR (GP_NATUREPIECEG="CFR"))';
   end else
   if (TypeTraitement = 'FACTURATION') then
   begin
      TCheckBox(GetControl('GP_VIVANTE')).Checked :=True;
      ecran.caption := TraduireMemoire ('Génération des factures fournisseurs manuelles');
      Naturepiece.Plus := ' AND ((GPP_NATUREPIECEG="CF") OR (GPP_NATUREPIECEG="CFR") OR (GPP_NATUREPIECEG="BLF") or (GPP_NATUREPIECEG="LFR") or (GPP_NATUREPIECEG="BFA"))';
      XX_WHERE.Text :='((GP_NATUREPIECEG="CF") OR (GP_NATUREPIECEG="CFR") OR (GP_NATUREPIECEG="BLF") or (GP_NATUREPIECEG="LFR") or (GP_NATUREPIECEG="BFA"))';
   end;
   if (TypeTraitement = 'RETOURFOU') then
   begin
      TCheckBox(GetControl('GP_VIVANTE')).Checked :=false;
      ecran.caption := TraduireMemoire ('Génération des Retours fournisseurs manuelles');
      Naturepiece.Plus := ' AND (GPP_NATUREPIECEG="ALV")';
      XX_WHERE.Text :='(GP_NATUREPIECEG="ALV")';
   end;
   UpdateCaption(Ecran);
end else if ecran.name = 'BTGROUPEMANNEG' then
begin
	 THdbGrid(GetCOntrol('FLISTE')).ondblclick := FlisteDblClick;
   Naturepiece := THMultiValComboBox (GetControl('GP_NATUREPIECEG'));
   XX_WHERE := THEdit(GetControl('XX_WHERE'));
   if (TypeTraitement = 'LIVRAISONCLI') then
   begin
      TCheckBox(GetControl('GP_VIVANTE')).Checked :=true;
      ecran.caption := TraduireMemoire ('Génération des livraisons regroupées');
      Naturepiece.plus := 'AND (GPP_NATUREPIECEG="CC")';
      XX_WHERE.Text  :='(GP_NATUREPIECEG="CC")';
   end;
   UpdateCaption(Ecran);

end;

{$ENDIF}
if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
  SetControlVisible ('TGP_AFFAIRE1', False);
CC:=THValComboBox(GetControl('GP_DOMAINE')) ; if CC<>Nil then PositionneDomaineUser(CC) ;
CC2:=THMultiValComboBox(GetControl('GP_ETABLISSEMENT')) ; if CC2<>Nil then PositionneEtabUser(CC2) ;
GLISTE:=THGRID(GetControl('GLISTE')) ;
BSELECTALL := TToolbarButton97(GetControl('bSelectAll'));
BSELECTALL.OnClick:=BSelectAllClick; BSELECTALL.Enabled:=False;
TOBSelect:=TOB.Create('_GenererPiece',NIL,-1);
Gliste.OnFlipSelection := FlipSelection;
end;

Procedure TOF_GCGeneManPiece.OnLoad ;
begin
inherited ;
  if Ecran.name <> 'BTGROUPEMANPIECE' then
  begin
    if (ctxAffaire in V_PGI.PGIContexte) And (GetControlText('GP_NATUREPIECEG')='CC') then
      SetControlText('XX_WHERE','GP_AFFAIRE=""');
  end;
  if GetControlVisible('GLISTE') then
  BEGIN
    SetControlVisible('GLISTE', False);
    SetControlVisible('FLISTE', True);
  END;
end;

Procedure TOF_GCGeneManPiece.OnUpdate ;
begin
inherited ;
if GLISTE.nbSelected > 0 then GLISTE.ClearSelected;
GLISTE.AllSelected:=False;
BSELECTALL.Down:=False; BSELECTALL.Enabled:=False;
SetControlEnabled('Bimprimer', True);
SetControlEnabled('BExport', True);
GLISTE.VidePile(False);
if TOBSelect.Detail.count > 0 then TOBSelect.ClearDetail;
InitGrille;
end;

Procedure TOF_GCGeneManPiece.OnClose ;
begin
inherited ;
TOBSelect.Free ; TOBSelect:=nil ;
end;

procedure TOF_GCGeneManPiece.FlipSelection(Sender: TObject);
var TOBL : TOB;
    TotalPieces : Double;
begin
  AffichageTotal;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_GCGeneManPiece.BSelectAllClick(Sender: TObject);
begin
GLISTE.ClearSelected ;
GLISTE.AllSelected:=BSELECTALL.Down ;
AffichageTotal;
end;

procedure TOF_GCGeneManPiece.InitGrille ;
var Col,Dec, FixedWidth, Larg : integer ;
    tal : TAlignment ;
    STitre,St,Ch,stA,FF,Typ,FPerso : string ;
    NomList,FRecordSource,FLien,FSortBy,FFieldList,FLargeur,FAlignement,FParams : string ;
    Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul,OkTri,OkNumCol : boolean ;
    Ftitre,tt,NC : Hstring;
begin
FixedWidth:=6;
GLISTE.ColCount:=60;
Col:=1; STitre:=' '; ColName:='Fixed';

With TFMul(Ecran) do
    BEGIN
    NomList:=Q.Liste;
    ChargeHListe(NomList,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC,FPerso,OkTri,OkNumCol);
    GLISTE.Titres.Add(FFieldList) ;
    While Ftitre<> '' do
        BEGIN
        StA:=ReadTokenSt(FAlignement);
        St:=ReadTokenSt(Ftitre);
        Ch:=ReadTokenSt(FFieldList);
        Larg:=ReadTokenI(FLargeur);
        tal:=TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
        if OkVisu then
            BEGIN
            GLISTE.Cells[Col,0]:=St ;
            GLISTE.ColAligns[Col]:=tal;
            GLISTE.ColWidths[Col]:=Larg*GLISTE.Canvas.TextWidth('W') ;;
            if OkLib then GLISTE.ColFormats[Col]:='CB=' + Get_Join(Ch);
            Typ:=ChampToType(Ch) ;
            if (Typ='INTEGER') or (Typ='SMALLINT') or (Typ='DOUBLE') then GLISTE.ColFormats[Col]:=FF ;
            STitre:=STitre+';'+St;
            ColName:=ColName+';'+Ch;
            inc (Col);
            END;
        END;
    GLISTE.ColCount:=Col ;
    AffecteGrid(GLISTE,taConsult) ;
    Hmtrad.ResizeGridColumns(GLISTE) ;
    GLISTE.ColWidths[0]:=FixedWidth;
    Hmtrad.ResizeGridColumns(GLISTE) ;
    END ;
end;

Procedure TOF_GCGeneManPiece.ChargeChampsRupt (NatPiece: string; var ChampsRupt : TStrings);
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT GPR_CONDGRP FROM PARPIECEGRP WHERE GPR_NATUREPIECEG="'+PieceGenere+
           '" AND GPR_NATPIECEGRP="' +NatPiece+ '"',True) ;
if Not Q.EOF then
   BEGIN
{$IFDEF EAGLCLIENT}
   ChampsRupt.Text:=Q.FindField('GPR_CONDGRP').AsString ;
{$ELSE}
   if Not TMemoField(Q.FindField('GPR_CONDGRP')).IsNull then
      ChampsRupt.Assign(TMemoField(Q.FindField('GPR_CONDGRP'))) ;
{$ENDIF}
   END else
   BEGIN
   with ChampsRupt do begin
    Add('PIECE;GP_TIERSFACTURE;=;');
    Add('PIECE;GP_DEVISE;=;');
    Add('PIECE;GP_TAUXDEV;=;');
{$IFNDEF BTP}
    Add('PIECE;GP_COTATION;=;');
{$ENDIF}
    Add('PIECE;GP_SAISIECONTRE;=;');
    Add('PIECE;GP_REGIMETAXE;=;');
    Add('PIECE;GP_ESCOMPTE;=;');
    Add('PIECE;GP_FACTUREHT;=;');
    if (THValComboBox(getControl('GP_VENTEACHAT'))<> nil) and
    	 (THValComboBox(getControl('GP_VENTEACHAT')).Value = 'VEN') then
    begin
    	Add('PIECE;GP_TIERSPAYEUR;=;');
    end;
    END;
   END;
// DBR Fiche 10665 - Debut
  if (GetInfoParPiece (PieceGenere, 'GPP_ECLATEAFFAIRE') = 'X') and (pos ('_AFFAIRE', ChampsRupt.Text) = 0) then
  begin
    ChampsRupt.Add ('PIECE;GP_AFFAIRE;=;');
  end;
  if (GetInfoParPiece (PieceGenere, 'GPP_ECLATEDOMAINE') = 'X') and (pos ('_DOMAINE', ChampsRupt.text) = 0) then
  begin
    ChampsRupt.Add ('PIECE;GP_DOMAINE;=;');
  end;
// DBR Fiche 10665 - Fin
Ferme(Q) ;
END ;

function TOF_GCGeneManPiece.AjustChampToStr(Champ : string; V1 : Variant) : string;
Var Typ, St : string;
begin
   Typ:=ChampToType(Champ) ; St:='' ;
   if (Typ='INTEGER') or (Typ='SMALLINT') then St:=IntToStr(VarAsType(V1,VArInteger)) else
    if (Typ='DOUBLE') or (Typ='RATE')  then St:=STRFPOINT(VarAsType(V1,VArDouble)) else
     if Typ='DATE'    then St:=USDateTime(VarAsType(V1,VArDate)) else
      if Typ='BOOLEAN' then St:=VarAsType(V1,VArString) else
        St:=VarAsType(V1,VarString) ;
Result:=St ;
end;

function TOF_GCGeneManPiece.FabriqueWherePiece (Tob_Piece : TOB) : string;
Var Where,Champ,Temp : String ;
    ChampsRupt : TStrings;
    CB : TCheckBox ;
    i_ind : integer;
    RemCur : double ;
    V1 : Variant;
begin
Where:='';
Where := RecupWhereCritere(TFmul(Ecran).Pages) ;
ChampsRupt:=TStringList.Create ;
ChargeChampsRupt(TOB_Piece.GetValue('GP_NATUREPIECEG'), ChampsRupt);
for i_ind:=0 to ChampsRupt.Count-1 do
    BEGIN
    Champ:=ChampsRupt[i_ind] ; Temp:=ReadTokenSt(Champ) ; Champ:=ReadTokenSt(Champ) ;
    if TOB_Piece.FieldExists(Champ) then
        BEGIN
        V1:=TOB_Piece.GetValue(Champ);
        Where:=Where+ ' AND ('+Champ+'="'+AjustChampToStr(Champ, V1)+'")';
        END ;
    END ;
ChampsRupt.Clear ; ChampsRupt.Free ;
CB:=TCheckBox(TFmul(Ecran).FindComponent('DEGROUPEREMISE')) ;
if CB.Checked then
    BEGIN
    RemCur:=TOB_Piece.GetValue('GP_REMISEPIED') ;
    Where:=Where + ' AND (GP_REMISEPIED=' + FloatToStr(RemCur) + ')';
    END;
Result:=Where;
end;

procedure TOF_GCGeneManPiece.ChargeLesPiecesRegroupables ;
var TOB_Piece :TOB;
    NaturepieceG ,Souche, Where : string;
    Numero, IndiceG : Integer ;
    QQ : Tquery ;
Begin
With TFMul(Ecran) do
begin
  {$IFDEF EAGLCLIENT}
  Q.TQ.Seek(FListe.Row-1) ;
  {$ENDIF}
  NaturepieceG := Q.FindField('GP_NATUREPIECEG').asstring ;
  Souche := Q.FindField('GP_SOUCHE').asstring ;
  Numero := Q.FindField('GP_NUMERO').asInteger ;
  IndiceG := Q.FindField('GP_INDICEG').asInteger ;
end;

Tob_Piece:=TOB.Create('PIECE', TOBSelect, -1) ;
QQ:=OpenSql('SELECT * From PIECE Where GP_NATUREPIECEG="'+NaturepieceG+'" AND GP_SOUCHE="'+Souche
     +'" AND GP_NUMERO='+IntToStr(Numero)+' AND GP_INDICEG='+IntToStr(IndiceG), TRUE) ;
Tob_Piece.SelectDB('',QQ);
ferme(QQ);
if TOB_Piece = nil then exit;
Where:=FabriqueWherePiece (Tob_Piece);
TOB_Piece.Free;
QQ := OpenSQL('SELECT * FROM PIECE ' + Where, True) ;
if not QQ.EOF then TOBSelect.LoadDetailDB('PIECE', '', '', QQ, FAlse);
ferme(QQ);
end;

function TOF_GCGeneManPiece.DemandeDatesLivraison(var DateLiv: TDateTime ) : boolean;
var TobDates : TOB;
begin
  TOBDates := TOB.Create ('LES DATES', nil,-1);
  TOBDates.AddChampSupValeur('RETOUROK','-');
  TOBDates.AddChampSupValeur('DATFAC',V_PGI.DateEntree);
  TOBDates.AddChampSupValeur('TYPEDATE','Date de livraison');
  TRY
    TheTOB := TOBDates;
    AGLLanceFiche('BTP','BTDEMANDEDATES','','','');
    TheTOB := nil;
    if TOBDates.getValue('RETOUROK')='X' then
    begin
    	DateLiv := TOBDates.GetDateTime('DATFAC');
    end;
  FINALLY
  	result := (TOBDates.getValue('RETOUROK')='X');
  	freeAndNil(TOBDates);
  END;
end;

procedure TOF_GCGeneManPiece.GenererPiece ;
var i_ind,Indice : integer;
    TOB_Piece,TOBMere : TOB ;
    CleDoc : R_CleDoc;
    St,Req : string;
    QQ : TQuery;
    OneDate : TDateTime;
begin
  With TFMul(Ecran) do
   BEGIN
    if Not GetControlVisible('GLISTE') then
        BEGIN
        BSELECTALL.Enabled:=True;
        SetControlEnabled('Bimprimer', False);
        SetControlEnabled('BExport', False);
        SetControlVisible('GLISTE', True);
        SetControlVisible('FLISTE', False);
        GLISTE.SetFocus;
        ChargeLesPiecesRegroupables;
        TOBSelect.PutGridDetail(GLISTE,False,False,ColName);
        for Indice := 0 to TOBselect.detail.count- 1 do
        begin
        	GListe.Objects[0,Indice+1]:=TOBselect.detail[Indice];
        end;
        AffichageTotal;
        GLISTE.ClearSelected ;
        GLISTE.AllSelected:=True;
        BSELECTALL.Down:=True;
        END else
        BEGIN
        if (GLISTE.NbSelected=0) then
            begin
            MessageAlerte('Aucun élément sélectionné');
            exit;
            end;
      if ecran.name = 'BTGROUPEMANNEG' then
      begin
       if (TypeTraitement = 'LIVRAISONCLI') then
       begin
         if not DemandeDatesLivraison(OneDate) then Exit;
       end;
      end;

        TOBMere:=TOB.Create('_GenererPiece',NIL,-1);
        for i_ind:=0 to GLISTE.NbSelected-1 do
            BEGIN
            GLISTE.GotoLeBOOKMARK(i_ind);
            {$IFDEF EAGLCLIENT}
            Q.TQ.Seek(FListe.Row-1) ;
            {$ENDIF}
            Tob_Piece:=TOB.Create('PIECE', TOBMere, -1) ;
        //					  ne marche plus si tri
        //            TOB_Piece.Dupliquer (TOBSelect.Detail[GLISTE.Row-1], False,True);
        //
						TOB_Piece.dupliquer ( TOB(Gliste.Objects[0,GListe.row]),false,true);
            END;
        if TOBMere.detail.count = 1 then
            BEGIN
            FillChar(CleDoc,Sizeof(CleDoc),#0) ;
            St:=TOBMere.Detail[0].GetValue('GP_NATUREPIECEG') + ';' +
                DateToStr (TOBMere.Detail[0].GetValue('GP_DATEPIECE')) + ';' +
	            TOBMere.Detail[0].GetValue('GP_SOUCHE') + ';' +
                IntToStr (TOBMere.Detail[0].GetValue('GP_NUMERO')) + ';' +
	            IntToStr (TOBMere.Detail[0].GetValue('GP_INDICEG')) + ';';
            TOBMere.Free ;
            StringToCleDoc(St,CleDoc) ;
            if CleDoc.NaturePiece<>'' then TransformePiece(CleDoc,PieceGenere,False) ;
            END else
            BEGIN
  			if ecran.name = 'BTGROUPEMANNEG' then
        begin
          RegroupePieces (TOBmere,'BLC',DateToStr(OneDate));
        end else
        begin
            CleDoc:=RegroupeLesPiecesMan(TOBMere,PieceRegroupe,False,False,True) ;
            if CleDoc.NaturePiece<>'' then TransformePiece(CleDoc,PieceGenere,False) ;
        end;
            TOBMere.Free ; // TOBMere:=nil ;
            END;
        BChercheClick(Nil);
        END ;
   END ;
end;

procedure TOF_GCGeneManPiece.AffichageTotal ;
var TotalPieces : Double;
    i, NbLignes : Integer;
    TOBL : TOB;
begin
  TotalPieces := 0;
  Nblignes := 0;
  for i := 1 to GListe.RowCount-1 do
    begin
    if Gliste.IsSelected(i) then
       begin
       TOBL := TOB(GListe.Objects [0,i]);
       TotalPieces := TotalPieces + TOBL.GetValue('GP_TOTALHTDEV');
       Inc(NbLignes);
       end;
    end;
  for i := 0 to TFMul(Ecran).PCumul.ControlCount - 1 do
  begin
    TFMul(Ecran).PCumul.Controls[i].visible := false;
  end;
  TFMul(Ecran).Pcumul.Caption := 'Total de la sélection ('+IntToStr(NbLignes)+' lignes) : '+  StrF00(TotalPieces,V_PGI.OkDecV);
end;

procedure TOF_GCGeneManPiece.VoirPiece ;
var TOB_Piece :TOB;
    NaturepieceG ,Souche, St : string;
    Numero, IndiceG : Integer ;
    DatePiece : TDateTime;
    CleDoc : R_CleDoc;
Begin
With TFMul(Ecran) do
   BEGIN
    if Not GetControlVisible('GLISTE') then
        BEGIN
        NaturepieceG:=Q.FindField('GP_NATUREPIECEG').asstring ;
        Souche:=Q.FindField('GP_SOUCHE').asstring ;
        Numero:=Q.FindField('GP_NUMERO').asInteger ;
        IndiceG:=Q.FindField('GP_INDICEG').asInteger ;
        DatePiece:=Q.FindField('GP_DATEPIECE').AsDateTime;
        END else
        BEGIN
        Tob_Piece:=TOB.Create('PIECE', Nil, -1) ;
        TOB_Piece.Dupliquer (TOBSelect.Detail[GLISTE.Row-1], False,True);
        NaturepieceG:=TOB_Piece.GetValue('GP_NATUREPIECEG') ;
        Souche:=TOB_Piece.GetValue('GP_SOUCHE') ;
        Numero:=TOB_Piece.GetValue('GP_NUMERO') ;
        IndiceG:=TOB_Piece.GetValue('GP_INDICEG') ;
        DatePiece:=TOB_Piece.GetValue('GP_DATEPIECE');
        TOB_Piece.Free ;
        END ;
    FillChar(CleDoc,Sizeof(CleDoc),#0) ;
    St:=NaturepieceG + ';' + DateToStr (DatePiece) + ';' + Souche + ';' +
        IntToStr (Numero) + ';' + IntToStr (IndiceG) + ';';
    StringToCleDoc(St,CleDoc) ;
    if CleDoc.NaturePiece<>'' then SaisiePiece(CleDoc,taConsult) ;
   END ;
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLGenererManPiece(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_GCGeneManPiece) then TOF_GCGeneManPiece(TOTOF).GenererPiece else exit;
end;

procedure AGLGeneManPiece_VoirPiece(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_GCGeneManPiece) then TOF_GCGeneManPiece(TOTOF).VoirPiece else exit;
end;

procedure TOF_GCGeneManPiece.FlisteDblClick(Sender: Tobject);
begin
	GenererPiece;
end;

Initialization
registerclasses([TOF_GCGeneManPiece]);
RegisterAglProc('GenererManPiece',TRUE,0,AGLGenererManPiece);
RegisterAglProc('GeneManPiece_VoirPiece',TRUE,0,AGLGeneManPiece_VoirPiece);
end.

