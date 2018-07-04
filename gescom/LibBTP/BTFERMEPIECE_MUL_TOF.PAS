{***********UNITE*************************************************
Auteur  ...... : L.S
Créé le ...... : 30/09/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTFERMEPIECE_MUL ()
Mots clefs ... : TOF;BTFERMEPIECE_MUL
*****************************************************************}
Unit BTFERMEPIECE_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
     Messages,
     AglInit,
     Ent1,
{$IFNDEF EAGLCLIENT}
		 fe_main,
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
		 MainEagl,
     eMul,Hqry,
{$ENDIF}
     M3Fp,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTofAfBaseCodeAffaire,
     UtilTOBPiece,
     HTB97,
     uTob,
     EntGc,
     RECALCPIECE_RAP_TOF,
     FactTob,
     uEntCommun ;

Type
  TOF_BTFERMEPIECE_MUL = Class (TOF_AFBASECODEAFFAIRE)
  private
  	TOBPieces : TOB;
    CBONLYCHANTIER : TcheckBox;
    procedure FermeCesPieces;
    procedure AddPieceAtraiter(Cledoc: r_cledoc; CodeAffaire,AffaireDevis : string);
    procedure CBONLYCHANTIERChange (Sender : TOBject);
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;

  end ;

Implementation

procedure TOF_BTFERMEPIECE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTFERMEPIECE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTFERMEPIECE_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTFERMEPIECE_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTFERMEPIECE_MUL.OnArgument (S : String ) ;
var CC : THValComboBox;
begin
	fMulDeTraitement := true;
  Inherited ;
  fTableName := 'PIECE';

  //Gestion Restriction Domaine et Etablissements
  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;
  if CC<>Nil then PositionneEtabUser(CC) ;

//  THValComboBox(GetControl('GP_NATUREPIECEG')).Plus := ' AND GPP_NATUREPIECEG IN ("ABT","FBT","LBT","CBT","DBT","ETU","BCE","BBO","BFC","DAC","FAC","PBT","FPR","FF","AF","AFS","BFC","CF","CFR","LFR","BLF")';
  THValComboBox(GetControl('GP_NATUREPIECEG')).Plus := ' AND GPP_NATUREPIECEG IN (SELECT GPP_NATUREPIECEG FROM PARPIECE WHERE GPP_MASQUERNATURE="-")';
  CBONLYCHANTIER := THCheckBox(GetCOntrol('CBONLYCHANTIER'));
  CBONLYCHANTIER.OnClick  := CBONLYCHANTIERChange;
  //
  TobPieces := Tob.create('LES PIECES', nil, -1);

end ;

procedure TOF_BTFERMEPIECE_MUL.OnClose ;
begin
  Inherited ;
  TOBPieces.free;
end ;

procedure TOF_BTFERMEPIECE_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTFERMEPIECE_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure AglBTFermePieces (parms:array of variant; nb: integer ) ;
var  F : TForm ;
     LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;
  if (LaTof is TOF_BTFERMEPIECE_MUL) then TOF_BTFERMEPIECE_MUL(LaTof).FermeCesPieces else exit;
end;

procedure TOF_BTFERMEPIECE_MUL.FermeCesPieces;
var
{$IFDEF EAGLCLIENT}
		Q : THquery;
{$ELSE}
		Q : Tquery;
{$ENDIF}
		F : TFmul;
    i : integer;
    Cledoc : r_cledoc;
    Ferme : boolean;
    Libferme : string;
    SRet : String;
    TOBPass : TOB;
begin
	Ferme := not (TCheckBox (GetCOntrol('GP_VIVANTE')).Checked );
  if Ferme then Libferme := 'réouvrir' else Libferme := 'fermer';
	if (PGIAsk ('Désirez-vous ' +LibFerme+' les pièces sélectionnées ?', ecran.Caption)<>mrYes) then exit;
  //
  TOBPass := TOB.create ('UNE TOB',nil,-1);
  TOBPASS.AddChampSupValeur ('RESULTAT','');
  TOBPASS.AddChampSupValeur ('DAYPASS','');
  TheTOB := TOBPAss;
 	Sret := AGLLAnceFiche('BTP','BTCONFIRMPASS','','','');
  TheTOB := nil;
  if TOBPass.GetValue('RESULTAT')<> 'OK' then
  begin
  	TOBPAss.free;
    exit;
  end;
  TOBPass.free;
  //
  F := TFmul(ecran);
  TRY
    if TFMul(F).Fliste.AllSelected then
    BEGIN
      Q:=TFmul(F).Q;
      Q.First;
      while Not Q.EOF do
      BEGIN
        FillChar(CleDoc, Sizeof(CleDoc), #0);
        cledoc.NaturePiece := Q.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche      := Q.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece := Q.FindField('GP_NUMERO').AsInteger;
        cledoc.DatePiece   := Q.FindField('GP_DATEPIECE').AsDateTime;
        cledoc.Indice      := Q.FindField('GP_INDICEG').AsInteger;
        AddPieceAtraiter (Cledoc,Q.FindField('GP_AFFAIRE').AsString,Q.FindField('GP_AFFAIREDEVIS').AsString);
        Q.NEXT;
      END;
    END
    ELSE
    BEGIN
      for i:=0 to TFMul(F).Fliste.nbSelected-1 do
      begin
        TFMul(F).Fliste.GotoLeBookmark(i);
        FillChar(CleDoc, Sizeof(CleDoc), #0);
{$IFDEF EAGLCLIENT}
      	TFMul(F).Q.TQ.Seek(TFMul(F).FListe.Row-1) ;

        cledoc.NaturePiece :=TFMul(F).Q.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche      :=TFMul(F).Q.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece :=TFMul(F).Q.FindField('GP_NUMERO').AsInteger;
        cledoc.DatePiece   :=TFMul(F).Q.FindField('GP_DATEPIECE').AsDateTime;
        cledoc.Indice      :=TFMul(F).Q.FindField('GP_INDICEG').AsInteger;
        AddPieceAtraiter (Cledoc,TFMul(F).Q.FindField('GP_AFFAIRE').AsString,TFMul(F).Q.FindField('GP_AFFAIREDEVIS').AsString);
{$ELSE}
        cledoc.NaturePiece :=TFMul(F).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche      :=TFMul(F).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece :=TFMul(F).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
        cledoc.DatePiece   :=TFMul(F).Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime;
        cledoc.Indice      :=TFMul(F).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
        AddPieceAtraiter (Cledoc,TFMul(F).Fliste.datasource.dataset.FindField('GP_AFFAIRE').AsString,
        												 TFMul(F).Fliste.datasource.dataset.FindField('GP_AFFAIREDEVIS').AsString);
{$ENDIF}
      end;
    END;
  finally
    if TOBPieces.detail.count > 0 then TraiteFermeturePieces (TOBPieces,Ferme);
    TFMul(F).Fliste.AllSelected:=False;
  	TOBPieces.ClearDetail;
  end;
end;

procedure TOF_BTFERMEPIECE_MUL.AddPieceAtraiter(Cledoc: r_cledoc; CodeAffaire,AffaireDevis : string);
var TOBPiece : TOB;
		QQ : Tquery;
begin
	TOBPIece := TOB.Create ('PIECE',TOBPieces,-1);
	TOBPIece.AddChampSupValeur('AFF_GENERAUTO','DIR');
  TOBPiece.putValue('GP_NATUREPIECEG',cledoc.NaturePiece);
  TOBPiece.putValue('GP_SOUCHE',cledoc.Souche);
  TOBPiece.putValue('GP_NUMERO',cledoc.NumeroPiece);
  TOBPiece.putValue('GP_INDICEG',cledoc.Indice);
  TOBPiece.putValue('GP_AFFAIRE',CodeAffaire);
  TOBPiece.putValue('GP_AFFAIREDEVIS',AffaireDevis);
  QQ := OpenSql('SELECT AFF_GENERAUTO FROM AFFAIRE WHERE AFF_AFFAIRE="'+AffaireDevis+'"',True,1,'',true);
  if not QQ.eof then
  begin
  	TOBPiece.putValue('AFF_GENERAUTO',QQ.FindField('AFF_GENERAUTO').AsString);
  end;
  Ferme (QQ);
end;

procedure TOF_BTFERMEPIECE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff:=THEdit(GetControl('GP_AFFAIRE'))   ;
  Aff0:=THEdit(GetControl('AFFAIRE0'));
  Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
  Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
  Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
  Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
  Tiers:=THEdit(GetControl('GP_TIERS'))   ;
end;


procedure TOF_BTFERMEPIECE_MUL.CBONLYCHANTIERChange(Sender: TOBject);
begin
	if CBONLYCHANTIER.State = cbgrayed then
  begin
  	THEDit(GetControl('XX_WHERE')).Text := '';
  end else if CBONLYCHANTIER.State = cbchecked then
  begin
  	THEDit(GetControl('XX_WHERE')).Text := ' AND GP_AFFAIRE<>''';
  end else
  begin
  	THEDit(GetControl('XX_WHERE')).Text := ' AND GP_AFFAIRE=''';
  end;
end;

Initialization
  registerclasses ( [ TOF_BTFERMEPIECE_MUL ] ) ;
	RegisterAglProc('BTFermeCesPieces', True , 0, AglBTFermePieces);
end.
