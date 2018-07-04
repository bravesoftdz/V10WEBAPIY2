{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 07/08/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRECALCPIECE_MUL ()
Mots clefs ... : TOF;BTRECALCPIECE_MUL
*****************************************************************}
Unit BTRECALCPIECE_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
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
     HTB97,
     uTob,
     EntGc,
     RECALCPIECE_RAP_TOF,
     FactTob,
     Ent1,uEntCommun,UtilTOBPiece ;

Type
  TOF_BTRECALCPIECE_MUL = Class (TOF_AFBASECODEAFFAIRE)
  private
  	CalculPv  : boolean;
  	TOBPieces : TOB;
    TOBLigOuv : TOB;
    procedure AddPieceAtraiter(Cledoc: r_cledoc);
    procedure RecalculeCesPieces;
    procedure AddLigneOuvrageAtraiter(Cledoc: R_CLedoc; TOBPiece: TOB);
    procedure Controlechamp(Champs, valeur: string);
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

procedure TOF_BTRECALCPIECE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECE_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECE_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECE_MUL.Controlechamp  (Champs , valeur : string);
begin
  if Champs = 'UNIQUEBLODBT' then
  begin
    THValComboBox(GetControl('GP_NATUREPIECEG')).Plus := ' AND GPP_NATUREPIECEG IN ("FBT","FBP","DAC")';
  end;
end;

procedure TOF_BTRECALCPIECE_MUL.OnArgument (S : String ) ;
var CC : THValComboBox;
    critere,champ,valeur : string;
    i : Integer;
begin
  Inherited ;

  TOBPieces := TOB.create ('LES PIECES',nil,-1);

  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;
  //
  //Gestion Restriction Domaine
  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  THValComboBox(GetControl('GP_NATUREPIECEG')).Plus := ' AND GPP_NATUREPIECEG IN ("ABT","FBT","LBT","CBT","DBT","ETU","BCE","BBO","BFC","DAC","FAC","PBT","CF","BLF","FF","LFR")';
  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;
end ;

procedure TOF_BTRECALCPIECE_MUL.OnClose ;
begin
  Inherited ;
  FreeAndNil(TOBPieces);
end ;

procedure TOF_BTRECALCPIECE_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECE_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECALCPIECE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1,
  Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff:=THEdit(GetControl('GP_AFFAIRE'))   ;
  Aff0:=THEdit(GetControl('GP_AFFAIRE0'));
  Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
  Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
  Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
  Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
  Tiers:=THEdit(GetControl('GP_TIERS'))   ;
end;

procedure TOF_BTRECALCPIECE_MUL.RecalculeCesPieces;
var Q : Tquery;
		F : TFmul;
    i : integer;
    Cledoc : r_cledoc;
    ReajustePVOuv : boolean;
begin
  ReajustePVOuv := false;
  F := TFmul(ecran);

  if F.Name='BTRECALCPIECE_MUL' then
  begin
  	CalculPv := (TCheckBox (GetCOntrol('CBCALCPV')).Checked );
  	ReajustePVOuv := (TCheckBox (GetCOntrol('CBREAJUSTEOUV')).Checked );
    if (PGIAsk ('Désirez-vous recalculer les pièces sélectionnés ?', ecran.Caption)<>mrYes) then exit;
  end
  else if F.Name='BTMANPIECE_MUL' then
  begin
  	CalculPv := false;
    if (PGIAsk ('Désirez-vous mettre à jour les documents sélectionnés ?', ecran.Caption)<>mrYes) then exit;
  end
  else
    Exit;

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
        AddPieceAtraiter (Cledoc);
        if F.Name='BTMANPIECE_MUL' then AddLigneOuvrageAtraiter (Cledoc, TOBPieces.detail[TOBPieces.Detail.count-1]);
        Q.NEXT;
      END;
      TFMul(F).Fliste.AllSelected:=False;
    END ELSE
    BEGIN
      for i:=0 to TFMul(F).Fliste.nbSelected-1 do
      begin
        TFMul(F).Fliste.GotoLeBookmark(i);
        FillChar(CleDoc, Sizeof(CleDoc), #0);
        cledoc.NaturePiece :=TFMul(F).Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
        Cledoc.Souche      :=TFMul(F).Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
        cledoc.NumeroPiece :=TFMul(F).Fliste.datasource.dataset.FindField('GP_NUMERO').AsInteger;
        cledoc.DatePiece   :=TFMul(F).Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime;
        cledoc.Indice      :=TFMul(F).Fliste.datasource.dataset.FindField('GP_INDICEG').AsInteger;
        AddPieceAtraiter (Cledoc);
        if F.Name='BTMANPIECE_MUL' then AddLigneOuvrageAtraiter (Cledoc, TOBPieces.detail[TOBPieces.Detail.count-1]);
      end;
    END;
  finally
    if F.Name='BTRECALCPIECE_MUL' then
    begin
      if TOBPieces.detail.count > 0 then TraiteRecalculPieces (TOBPieces,CalculPv,ReajustePVOuv);
    end
    else if F.Name='BTMANPIECE_MUL' then
    begin
      if TOBPieces.detail.count > 0 then TraiteManPieces (TOBPieces);
    end;
  	TOBPieces.ClearDetail;
  end;
end;

procedure TOF_BTRECALCPIECE_MUL.AddPieceAtraiter(Cledoc: r_cledoc);
var TOBPiece : TOB;
		QQ : Tquery;
begin
	TOBPIece := TOB.Create ('PIECE',TOBPieces,-1);
  QQ := OpenSql ('SELECT * FROM PIECE WHERE '+WherePiece (cledoc,ttdPiece,true),True);
  TOBPiece.SelectDB ('',QQ);
  Ferme (QQ);
end;

procedure TOF_BTRECALCPIECE_MUL.AddLigneOuvrageAtraiter(Cledoc : R_CLedoc; TOBPiece : TOB);
var StSQL       : String;
    QQ          : TQuery;
begin

  StSQL := 'SELECT * FROM LIGNEOUV WHERE ' + WherePiece (cledoc,ttdOuvrage,true);
  StSQL := StSQl + 'ORDER BY BLO_NUMLIGNE, BLO_N1, BLO_N2, BLO_N3, BLO_N4, BLO_N5';

  QQ := OpenSql(StSQL, True) ;

  TOBPiece.LoadDetailDB  ('LIGNEOUV','','',QQ,false);

  Ferme (QQ);

end;

procedure AglBTRecalcPieces (parms:array of variant; nb: integer ) ;
var  F : TForm ;
     LaTof : TOF;
begin

  F:=TForm(Longint(Parms[0])) ;

  if (F is TFMul) then Latof:=TFMul(F).Latof else exit;

  if (LaTof is TOF_BTRECALCPIECE_MUL) then
    TOF_BTRECALCPIECE_MUL(LaTof).RecalculeCesPieces
  else
    exit;

end;

Initialization
  registerclasses ( [ TOF_BTRECALCPIECE_MUL ] ) ;
	RegisterAglProc('RecalculeCesPieces', True , 0, AglBTRecalcPieces);
end.

