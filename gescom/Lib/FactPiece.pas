unit FactPiece ;

interface

uses HEnt1, UTOB,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls, FactTOB, SysUtils, EntGC , UentCommun,Hmsgbox;

//Procedure NumeroteLignesGC (GS: THGrid; TOBPiece: TOB; ForceRenum: Boolean = False); { NEWPIECE }
Procedure NumeroteLignesGC (GS: THGrid; TOBPiece: TOB; ForceRenum: Boolean = False; GestionAffichage: boolean = True); { NEWPIECE }
procedure EvaluerMaxNumOrdre(TOBPiece: TOB);
procedure MiseAJourGrid(GS: THGrid; TOBPiece: TOB; CodeArticle: string; Qte: Double; ARow: Integer);
function  SupDimAZero(GS: THGrid; CodeArticle: string; TOBDim, TOBPiece: TOB; ARow: Integer): Integer;
function  GetSoucheG(NaturePieceG, Etablissement, Domaine: String3): String3;
function  GetNumSoucheG(SoucheG: String3; ODate : TdateTime): integer;
procedure IncNumSoucheG(SoucheG: String3; ODate : TdateTime);
function  HistoPiece(TobPiece_O: Tob): Boolean;
function  SetNumberAttribution(Nature,SoucheG: string; DatePiece : TDateTime; NbPiece: integer): integer;
function  SetDefinitiveNumber(TOBPiece, TOBBases, TOBBasesL,TOBEches, TOBNomenclature, TOBAcomptes, TOBPieceRG, TOBBasesRG, TobLigneTarif: TOB; NumDef: integer): boolean;
function  SetNumeroDefinitif(TOBPiece, TOBBases, TOBBasesL,TOBEches, TOBNomenclature, TOBAcomptes, TOBPIeceRG, TOBBasesRG, TobLigneTarif: TOB): boolean;
function  IsTransfert(NaturePiece : String): Boolean;
function LireMaxNumOrdre(TOBPiece: TOB): integer;
procedure EcrireMaxNumordre(TOBPiece: TOB; MaxNumOrdre: integer);
function GetMaxNumOrdre (TOBpiece : TOB) : integer;
procedure JustNumerote (TOBPiece : TOb; Lignedepart: integer=-1);
procedure SuprimeParagraphesOrphelin(TobFacture: TOB);
function IsDetailInside(IndDep: integer; TobFacture,TOBL: TOB): boolean;
procedure SuprimeThisParagraphe(TobFacture,TOBL : TOB; var IndDep: integer);
procedure ReajusteNumOrdre (TOBPiece,TOBPieceRg : TOB) ;
  //FV1
  Procedure SupprimeOuvrageVide(TobIntermediaire : TOB);
procedure StockeInfoTypeLigne (TOBL : TOB; ValLoc : T_Valeurs);

implementation

uses
  FactTarifs,
  FactUtil,
  UtilTOBPiece
  ;
procedure EcrireMaxNumordre(TOBPiece: TOB; MaxNumOrdre: integer);
begin
  TOBPiece.PutValue('GP_CODEORDRE', MaxNumOrdre);
end;

procedure ReajusteNumOrdre (TOBPiece,TOBPieceRg : TOB) ;
var II : integer;
    MaxNumOrdre : integer;
begin
  maxNumordre:= LireMaxNumOrdre(TOBpiece);
  if TOBpieceRg.detail.count > 0 then
  begin
    for II := 0 to TOBpieceRG.detail.count -1 do
    begin
      if TOBpieceRG.detail[II].FieldExists ('NUMORDRE') then
      begin
        if TOBpieceRG.detail[II].getInteger ('NUMORDRE')> MaxNumOrdre then
        begin
          MaxNumOrdre := TOBpieceRG.detail[II].getInteger ('NUMORDRE');
        end;
      end;
    end;
    EcrireMaxNumordre (TOBpiece,MaxNumOrdre);
  end;
end;

procedure EvaluerMaxNumOrdre(TOBPiece: TOB);

  function GetMaxUniqueBlo (TOBPiece : TOB) : integer;
  var QQ : TQuery;
      cledoc : r_cledoc;
  begin
    cledoc := TOB2CleDoc(TOBpiece);
    result := 1;
    QQ := OpenSql ('SELECT MAX(BLO_UNIQUEBLO) FROM LIGNEOUV WHERE '+WherePiece(cledoc,ttdOuvrage,False),True,1,'',True);
    if not QQ.Eof then
    begin
      Result := QQ.fields[0].AsInteger+1;
    end;
    Ferme(QQ);
  end;

var i, NumOrdre, MaxNumOrdre, iNO: integer;
  TOBL: TOB;
begin
  iNO := 0;
  MaxNumOrdre := 0;
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    if i = 0 then iNO := TOBL.GetNumChamp('GL_NUMORDRE');
    NumOrdre := TOBL.GetValeur(iNO);
    if NumOrdre > MaxNumOrdre then MaxNumOrdre := NumOrdre;
  end;
  EcrireMaxNumOrdre(TOBPiece, MaxNumOrdre);
  TOBPiece.SetInteger('GP_UNIQUEBLO',GetMaxUniqueBlo(TOBpiece));
end;

function LireMaxNumOrdre(TOBPiece: TOB): integer;
begin
//  if TOBPiece.GetValue('GP_CODEORDRE') <= 0 then EvaluerMaxNumOrdre(TOBPiece);
  Result := TOBPiece.GetValue('GP_CODEORDRE');
end;


function GetMaxNumOrdre (TOBpiece : TOB) : integer;
var MaxNumOrdre : integer;
begin
  MaxNumOrdre:=LireMaxNumOrdre(TOBPiece);
  Inc(MaxNumOrdre);
  EcrireMaxNumOrdre(TOBPiece, MaxNumOrdre) ;
  result := MaxNumOrdre;
end;

procedure JustNumerote (TOBPiece : TOb; Lignedepart: integer=-1);
var Depart,Inl,i : integer;
    TOBL : TOB;
begin
  if TOBPiece.detail.count >0 then
  begin
    Inl := TOBPiece.detail[0].GetNumChamp('GL_NUMLIGNE') ;
    if LigneDepart <> -1 Then depart := LigneDepart else depart := 0;
    for i := depart to TOBPiece.Detail.Count-1 do
    begin
      TOBL := GetTOBLigne(TOBPiece, i+1);
      if TOBL <> nil then TOBL.PutValeur(Inl, i+1);
    end;
  end;
end;

//Procedure NumeroteLignesGC(GS: THGrid; TOBPiece: TOB; ForceRenum: Boolean = False); { NEWPIECE }
Procedure NumeroteLignesGC (GS: THGrid; TOBPiece: TOB; ForceRenum: Boolean = False; GestionAffichage: boolean = True); { NEWPIECE }
var
  i, inl, ino, NumOrdre, MaxNumOrdre: integer;
  TOBL: TOB;
  TobLigneTarif: Tob;
begin
  inl:=0; ino:=0;
  MaxNumOrdre:=LireMaxNumOrdre(TOBPiece);
  for i := 0 to TOBPiece.Detail.Count-1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, i+1);
    if i=0 then
    begin
      inl := TOBL.GetNumChamp('GL_NUMLIGNE') ;
      ino := TOBL.GetNumChamp('GL_NUMORDRE') ;
    end;
    // Modif BTP
    if GestionAffichage then
      if GS <> nil then GS.Cells[SG_NL,i+1]:=IntToStr(i+1) ;
    // --
    TOBL.PutValeur(inl, i+1);
    NumOrdre := TOBL.GetValeur(ino);
    if (NumOrdre<=0) or ForceRenum then { NEWPIECE }
    begin
      Inc(MaxNumOrdre);
      TOBL.PutValeur(ino, MaxNumOrdre);
    end;
  end;
  EcrireMaxNumOrdre(TOBPiece, MaxNumOrdre) ;
  TobLigneTarif := GetTobLigneTarifInTobPiece(TobPiece);
  NumeroteLigneTarif(TobLigneTarif);
end;

procedure MiseAJourGrid(GS: THGrid; TOBPiece: TOB; CodeArticle: string; Qte: Double; ARow: Integer);
var TOBL: TOB;
  Ligne, i: Integer;
  TotQte: Double;
begin
  i := ARow;
  TOBL := TOBPiece.Detail[i - 1];
  if TOBL = nil then Exit;
  while TOBL.GetValue('GL_CODESDIM') <> CodeArticle do
  begin
    Dec(i);
    if i >= 1 then TOBL := TOBPiece.Detail[i - 1] else
    begin
      Break;
      TOBL := nil;
    end;
    {JLD ! Attention pas de break provoquait boucle infinie !!!!}
  end;
  if TOBL = nil then Exit;
  TotQte := TOBL.GetValue('GL_QTEFACT');
  Ligne := TOBL.GetValue('GL_NUMLIGNE');
  if TOBL.GetValue('GL_TYPEDIM') = 'GEN' then
  begin
    TOBL.PutValue('GL_QTEFACT', TotQte - Qte);
    TOBL.PutValue('GL_QTESTOCK', TotQte - Qte);
    if TOBL.Getvalue('GL_QTEFACT') = 0 then // sup de la ligne GEN si plus de DIM (QTEDIM=0)
    begin
      GS.DeleteRow(Ligne);
      TOBPiece.Detail[Ligne - 1].Free;
      DeleteTobLigneTarif(TobPiece, Ligne - 1);
      NumeroteLignesGC(GS, TOBPiece);
    end;
  end;
end;

function SupDimAZero(GS: THGrid; CodeArticle: string; TOBDim, TOBPiece: TOB; ARow: Integer): Integer;
var NbSup, i, Ligne, Qte: Integer;
  TOBL, TOBItem: TOB;
begin
  NbSup := 0;
  i := ARow;
  Qte := 0;
  TOBL := TOBPiece.Detail[i];
  while TOBL.GetValue('GL_TYPEDIM') = 'DIM' do
  begin
    TOBItem := TOBDim.FindFirst(['GA_ARTICLE'], [TOBL.GetValue('GL_ARTICLE')], False);
    if TOBItem = nil then
    begin
      Ligne := TOBL.GetValue('GL_NUMLIGNE');
      Qte := Qte + TOBL.GetValue('GL_QTEFACT');
      GS.DeleteRow(Ligne);
      TOBPiece.Detail[Ligne - 1].Free;
      DeleteTobLigneTarif(TobPiece, Ligne - 1);
      NbSup := NbSup + 1;
      NumeroteLignesGC(GS, TOBPiece);
    end else i := i + 1;
    if i > TOBPiece.Detail.Count - 1 then break;
    TOBL := TOBPiece.Detail[i];
  end;
  Result := NbSup;
  if TOBDim.Detail.Count = 0 then
  begin
    MiseAJourGrid(GS, TOBPiece, CodeArticle, Qte, ARow);
  end;
end;

function GetSoucheG(NaturePieceG, Etablissement, Domaine: String3): String3;
begin
  {$IFDEF NOMADE}
  Result := GetColonneSQL('SITEPIECE', 'GSP_SOUCHE', 'GSP_NATUREPIECEG="' + NaturePieceG + '" and GSP_CODESITE="' + PCP_LesSites.LeSiteLocal.SSI_CODESITE + '"');
  if (result = '') and (V_PGI.SuperViseur) then
  begin
  {$ENDIF}
    result := '';
    if Etablissement <> '' then Result := GetInfoParPieceCompl(NaturePieceG, Etablissement, 'GPC_SOUCHE');
    if (Result = '') and (Domaine <> '') then Result := GetInfoParPieceDomaine(NaturePieceG, Domaine, 'GDP_SOUCHE');
    if Result = '' then Result := GetInfoParPiece(NaturePieceG, 'GPP_SOUCHE');
  {$IFDEF NOMADE}
  end;
  {$ENDIF}
end;

procedure AddCptMensuel (Souche: string; Year,Month : Word);
var TOBX : TOB;
begin
	TOBX := TOB.Create ('BCPTMENSUEL',nil,-1);
  TOBX.PutValue('BCM_NATUREPIECEG',Souche);
  TOBX.PutValue('BCM_ANNEE',IntToStr(Year));
  TOBX.PutValue('BCM_MOIS',IntToStr(Month));
  TOBX.PutValue('BCM_COMPTEUR',1);
  TOBX.InsertDB(nil);
  TOBX.free;
end;


function GetNumSoucheG(SoucheG: String3; ODate : TdateTime): integer;
var Q: TQuery;
		Year,Month,Day : word;
    req ,cpt : string;

begin
  Result := 0;
  if SoucheG = '' then Exit;
  //FV1 - 23/11/2015 : FS#1688 - SERVAIS : Pb en paramétrage de compteur mensuel.
  //if GetInfoParPiece(SoucheG, 'GPP_NUMMOISPIECE') = 'X' then
  if GetInfoParSouche(SoucheG, 'BS0_NUMMOISPIECE') = 'X' then
  begin
  	decodedate(ODate, Year, Month, Day);
  	req := 'SELECT BCM_COMPTEUR FROM BCPTMENSUEL WHERE  BCM_NATUREPIECEG="' + SoucheG + '" '+
    			 'AND BCM_ANNEE=' + IntToStr(Year) + ' AND '+
           'BCM_MOIS=' + IntToStr(Month);
    if not ExisteSQL(Req) then
    begin
      AddCptMensuel (SoucheG,Year,Month);
    end;
    Q := OpenSQL(Req,True,1,'',true);
    if not Q.EOF then Result := Q.Fields[0].AsInteger;
    Cpt  := copy(IntToStr(Year),3,2) + StringOfChar('0', 2 - Length(IntToStr(Month))) + IntToStr(Month) + StringOfChar('0', 4 - Length(IntToStr(Result))) + IntToStr(Result);   //N° = 4 de Long
    Result    := StrToInt(cpt);
  	Ferme(Q);
  end else
  begin
    Q := OpenSQL('SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="' + SoucheG + '"', True,-1, '', True);
    if not Q.EOF then Result := Q.Fields[0].AsInteger;
  	Ferme(Q);
  end;
end;

function GetNumPieceG(NaturePieceG, Etablissement, Domaine: String3; ODate : TdateTime): integer;
var SoucheG: String3;
begin
  Result := 0;
  SoucheG := GetSoucheG(NaturePieceG, Etablissement, Domaine);
  if SoucheG = '' then Exit;
  Result := GetNumSoucheG(SoucheG,ODate);
end;

procedure IncNumSoucheG(SoucheG: String3; ODate : TdateTime);
var Year,Month,Day : word;
begin
  if SoucheG = '' then Exit;

  //FV1 - 23/11/2015 : FS#1688 - SERVAIS : Pb en paramétrage de compteur mensuel.
  //if GetInfoParPiece(SoucheG, 'GPP_NUMMOISPIECE') = 'X' then
  if GetInfoParSouche(SoucheG, 'BS0_NUMMOISPIECE') = 'X' then
  begin
  	decodedate(ODate, Year, Month, Day);
    ExecuteSQL('UPDATE BCPTMENSUEL SET BCM_COMPTEUR=BCM_COMPTEUR +1  WHERE  BCM_NATUREPIECEG="' + SoucheG + '" AND BCM_ANNEE=' + IntToStr(Year) + ' AND BCM_MOIS=' + IntToStr(Month));
  end else
  begin
  	ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART=SH_NUMDEPART+1 WHERE SH_TYPE="GES" AND SH_SOUCHE="' + SoucheG + '"');
  end;
end;

procedure IncNumPieceG(NaturePieceG, Etablissement, Domaine: String3; ODate : TdateTime);
var SoucheG: String3;
begin
  if NaturePieceG = '' then Exit;
  SoucheG := GetSoucheG(NaturePieceG, Etablissement, Domaine);
  if SoucheG = '' then Exit;
  IncNumSoucheG(SoucheG,Odate);
end;

function HistoPiece(TobPiece_O: Tob): Boolean; { NEWPIECE }
begin
  if Assigned(TobPiece_O) then
    Result := (GetInfoParPiece(TOBPiece_O.GetValue('GP_NATUREPIECEG'), 'GPP_HISTORIQUE') = 'X')
  else
    Result := False;
end;

function SetNumberAttribution(Nature,SoucheG: string; DatePiece : TDateTime; NbPiece: integer): Integer;
var QQ      : TQuery;
  NumDef    : Longint;
  Nb, ii    : integer;
  Okok      : boolean;
  //FV1 : Intégration Numérotation mensuelle
  Month     : word;
  Year      : word;
  Day       : word;
  NumPiece  : String;
  ok_Cpt    : Boolean;
//  NbCar     : Integer;
  StSQL     : String;
begin
  Result := 0;

  if SoucheG = '' then Exit;

  if NbPiece <= 0 then NbPiece := 1;

  decodedate(DatePiece, Year, Month, Day);
  //FV1 - 23/11/2015 : FS#1688 - SERVAIS : Pb en paramétrage de compteur mensuel.
  //if GetInfoParPiece(Nature, 'GPP_NUMMOISPIECE') = 'X' then Ok_Cpt := True else Ok_Cpt := False;
  if GetInfoParSouche(SoucheG, 'BS0_NUMMOISPIECE') = 'X' then Ok_Cpt := True else Ok_Cpt := False;

  ii := 0;
  repeat
    NumDef := -1;
    inc(ii);
    if Ok_Cpt then
    begin
      StSQL := 'SELECT BCM_COMPTEUR FROM BCPTMENSUEL WHERE  BCM_NATUREPIECEG="' + SoucheG + '" '+
             'AND BCM_ANNEE=' + IntToStr(Year) + ' AND '+
             'BCM_MOIS=' + IntToStr(Month);
      if not ExisteSQL(StSQL) then
      begin
        AddCptMensuel (SoucheG,Year,Month);
      end;
      StSQL := 'SELECT BCM_COMPTEUR FROM BCPTMENSUEL WHERE BCM_NATUREPIECEG="' + SoucheG + '" AND BCM_ANNEE=' + IntToStr(Year) + ' AND BCM_MOIS=' + IntToStr(Month)
    end else
    begin
      StSQL := 'SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="' + SoucheG + '"';
    end;
    QQ := OpenSQL(StSQL, True,-1, '', True);
    if not QQ.EOF then NumDef := QQ.Fields[0].AsInteger;
    Ferme(QQ);
    if NumDef <= 0 then exit;
    if Ok_Cpt then
      StSQL := 'UPDATE BCPTMENSUEL SET BCM_COMPTEUR=' + IntToStr(NumDef + NbPiece) + ' WHERE  BCM_NATUREPIECEG="' + SoucheG + '" AND BCM_ANNEE=' + IntToStr(Year) + ' AND BCM_MOIS=' + IntToStr(Month) + ' AND BCM_COMPTEUR=' + IntToStr(NumDef)
    else
      StSQL := 'UPDATE SOUCHE SET SH_NUMDEPART=' + IntToStr(NumDef + NbPiece) + ' WHERE SH_TYPE="GES" AND SH_SOUCHE="' + SoucheG + '" AND SH_NUMDEPART=' + IntToStr(NumDef);
    Nb := ExecuteSQL(StSQL);
    Okok := (Nb = 1);
  until ((Okok) or (ii > 20));
  if not Okok then
  begin
    PgiError('Erreur Attribution de numéro de pièce');
    V_PGI.IoError := oeUnknown;
    Exit;
  end;

  if Ok_Cpt then
  begin
    Numpiece  := copy(IntToStr(Year),3,2) + StringOfChar('0', 2 - Length(IntToStr(Month))) + IntToStr(Month) + StringOfChar('0', 4 - Length(IntToStr(NumDef))) + IntToStr(NumDef);   //N° = 4 de Long
    Result    := StrToInt(Numpiece);
  end
  else
    Result := NumDef;

end;

function SetDefinitiveNumber(TOBPiece, TOBBases, TOBBasesL,TOBEches, TOBNomenclature, TOBAcomptes, TOBPieceRG, TOBBasesRG, TobLigneTarif: TOB; NumDef: integer): boolean;
var SoucheG: String3;
  i, iNum: integer;
  j: integer;
  TOBB: TOB;
begin
  Result := False;
  iNum := 0;
  SoucheG := TOBPiece.GetValue('GP_SOUCHE');
  if SoucheG = '' then
  begin
    PgiError('Erreur Souche de la pièce indéfinie');
    Exit;
  end;
  if NumDef <= 0 then
  begin
    PgiError('Erreur numéro de pièce indéfinie');
    exit;
  end;
  Result := True;
  if NumDef <> TOBPiece.GetValue('GP_NUMERO') then
  begin
    TOBPiece.PutValue('GP_NUMERO', NumDef);
    for i := 0 to TOBPiece.Detail.Count - 1 do
    begin
      if i = 0 then iNum := TOBPiece.Detail[i].GetNumChamp('GL_NUMERO');
      TOBB := TOBPiece.Detail[i];
      TOBB.PutValeur(iNum, NumDef);
    end;
    if TOBBases <> nil then for i := 0 to TOBBases.Detail.Count - 1 do
      begin
        if i = 0 then iNum := TOBBases.Detail[i].GetNumChamp('GPB_NUMERO');
        TOBB := TOBBases.Detail[i];
        TOBB.PutValeur(iNum, NumDef);
      end;
    if TOBBasesL <> nil then for i := 0 to TOBBasesL.Detail.Count - 1 do
      begin
        if i = 0 then iNum := TOBBasesL.Detail[i].GetNumChamp('BLB_NUMERO');
        TOBB := TOBBasesL.Detail[i];
        TOBB.PutValeur(iNum, NumDef);
      end;
    if TOBEches <> nil then for i := 0 to TOBEches.Detail.Count - 1 do
      begin
        if i = 0 then iNum := TOBEches.Detail[i].GetNumChamp('GPE_NUMERO');
        TOBB := TOBEches.Detail[i];
        TOBB.PutValeur(iNum, NumDef);
      end;
    if TOBAcomptes <> nil then
    begin
      TOBAcomptes.AddChampSup('CHGTNUM',False); // JT eQualité 10246 (Pour maj ECRITURE si chgt numéro)
      for i := 0 to TOBAcomptes.Detail.Count - 1 do
        begin
          if i = 0 then iNum := TOBAcomptes.Detail[i].GetNumChamp('GAC_NUMERO');
          TOBB := TOBAcomptes.Detail[i];
          TOBB.PutValeur(iNum, NumDef);
          // JT Modif libellé acomptes
          TOBB.PutValue('GAC_LIBELLE',Copy('Ac./Règ. ' + DateToStr(TOBPiece.GetValue('GP_DATEPIECE')) + ' Pce. '+ IntToStr(NumDef), 1, 35));
        end;
    end;
    if TOBPieceRG <> nil then for i := 0 to TOBPieceRG.Detail.Count - 1 do
      begin
        if i = 0 then iNum := TOBPIECERg.Detail[i].GetNumChamp('PRG_NUMERO');
        TOBB := TOBPIECERG.Detail[i];
        TOBB.PutValeur(iNum, NumDef);
      end;
    if TOBBasesRg <> nil then for i := 0 to TOBBasesRg.Detail.Count - 1 do
      begin
        if i = 0 then iNum := TOBBasesRg.Detail[i].GetNumChamp('PBR_NUMERO');
        TOBB := TOBBasesRG.Detail[i];
        TOBB.PutValeur(iNum, NumDef);
      end;

    if Assigned(TobLigneTarif) then
    begin
      for i := 0 to TobLigneTarif.Detail.Count - 1 do
      begin
        if i = 0 then iNum := TobLigneTarif.Detail[i].GetNumChamp('GL_NUMERO');
        TobB := TobLigneTarif.Detail[i];
        TobB.PutValeur(iNum, NumDef);
        for j := 0 to TobB.Detail.Count - 1 do TobB.Detail[j].PutValue('GLT_NUMERO', NumDef)
      end;
    end;
  end;
end;

function SetNumeroDefinitif(TobPiece, TobBases, TOBBasesL,TobEches, TobNomenclature, TobAcomptes, TobPIeceRG, TobBasesRG, TobLigneTarif : tob): boolean;
var SoucheG: String3;
//  QQ: TQuery;
  NumDef: Longint;
  {Nb,} i, iNum: integer;
  j           : integer;
  TOBB: TOB;
begin
  Result := False;
//  NumDef := 0;
  iNum := 0;
  SoucheG := TOBPiece.GetValue('GP_SOUCHE');
  if SoucheG = '' then Exit;
  (*
  QQ := OpenSQL('SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="' + SoucheG + '"', True,-1, '', True);
  if not QQ.EOF then NumDef := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  *)
  NumDef := SetNumberAttribution(TOBPiece.GetValue('GP_NATUREPIECEG'),SoucheG,TOBPiece.GetValue('GP_DATEPIECE'),1);
  if NumDef > 0 then
  begin
    (*
    Nb := ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART=' + IntToStr(NumDef + 1) + ' WHERE SH_TYPE="GES" AND SH_SOUCHE="' + SoucheG + '"');
    if Nb <> 1 then
    begin
      V_PGI.IoError := oeUnknown;
      Exit;
    end;
    *)
    Result := True;
    if NumDef <> TOBPiece.GetValue('GP_NUMERO') then
    begin
      TOBPiece.PutValue('GP_NUMERO', NumDef);
      for i := 0 to TOBPiece.Detail.Count - 1 do
      begin
        if i = 0 then iNum := TOBPiece.Detail[i].GetNumChamp('GL_NUMERO');
        TOBB := TOBPiece.Detail[i];
        TOBB.PutValeur(iNum, NumDef);
      end;
      if TOBBases <> nil then for i := 0 to TOBBases.Detail.Count - 1 do
        begin
          if i = 0 then iNum := TOBBases.Detail[i].GetNumChamp('GPB_NUMERO');
          TOBB := TOBBases.Detail[i];
          TOBB.PutValeur(iNum, NumDef);
        end;
      if TOBBasesL <> nil then for i := 0 to TOBBasesL.Detail.Count - 1 do
        begin
          if i = 0 then iNum := TOBBasesL.Detail[i].GetNumChamp('BLB_NUMERO');
          TOBB := TOBBasesL.Detail[i];
          TOBB.PutValeur(iNum, NumDef);
        end;
      if TOBEches <> nil then for i := 0 to TOBEches.Detail.Count - 1 do
        begin
          if i = 0 then iNum := TOBEches.Detail[i].GetNumChamp('GPE_NUMERO');
          TOBB := TOBEches.Detail[i];
          TOBB.PutValeur(iNum, NumDef);
        end;
      if TOBAcomptes <> nil then for i := 0 to TOBAcomptes.Detail.Count - 1 do
        begin
          if i = 0 then iNum := TOBAcomptes.Detail[i].GetNumChamp('GAC_NUMERO');
          TOBB := TOBAcomptes.Detail[i];
          TOBB.PutValeur(iNum, NumDef);
        end;
      if TOBPIECERG <> nil then for i := 0 to TOBPIECERG.Detail.Count - 1 do
        begin
          if i = 0 then iNum := TOBPIECERG.Detail[i].GetNumChamp('PRG_NUMERO');
          TOBB := TOBPIECERG.Detail[i];
          TOBB.PutValeur(iNum, NumDef);
        end;
      if TOBBasesRG <> nil then for i := 0 to TOBBasesRG.Detail.Count - 1 do
        begin
          if i = 0 then iNum := TOBBasesRG.Detail[i].GetNumChamp('PBR_NUMERO');
          TOBB := TOBBasesRG.Detail[i];
          TOBB.PutValeur(iNum, NumDef);
        end;

      if Assigned(TobLigneTarif) then
      begin                                                                                       
        for i := 0 to TobLigneTarif.Detail.Count - 1 do                                           
        begin                                                                                     
          if i = 0 then iNum := TobLigneTarif.Detail[i].GetNumChamp('GL_NUMERO');                 
          TobB := TobLigneTarif.Detail[i];                                                        
          TobB.PutValeur(iNum, NumDef);                                                           
          for j := 0 to TobB.Detail.Count - 1 do TobB.Detail[j].PutValue('GLT_NUMERO', NumDef)    
        end;                                                                                      
      end;                                                                                        
    end;
  end else
  begin
    V_PGI.IoError := oeUnknown;
    Exit;
  end;

end;

function IsTransfert(NaturePiece : String): Boolean;
begin
  if (NaturePiece = 'TEM') or (NaturePiece = 'TRV') or (NaturePiece = 'TRE') then
    Result := True
  else Result := False;
end;

procedure SuprimeParagraphesOrphelin(TobFacture: TOB);
var Indice : integer;
    TOBL : TOB;
begin
  Indice := TOBFacture.detail.count -1;
  repeat
    if Indice >= 0 then
    begin
      TOBL := TOBFacture.detail[Indice];
      if IsFinParagraphe(TOBL) then
      begin
        if not IsDetailInside (Indice,TobFacture, TOBL) then
        begin
          SuprimeThisParagraphe (TobFacture,TOBL,Indice);
        end;
      end;
      dec(Indice);
    end;
  until indice <= 0;
end;

function IsDetailInside(IndDep: integer; TobFacture,TOBL: TOB): boolean;
var indice : integer;
    Niveau : integer;
    TOBI : TOB;
begin
  result := false;
  Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
  for Indice := Inddep-1 downto 0 do
  begin
    TOBI := TOBFacture.detail[Indice];
    if IsArticle (TOBI) then BEGIN result := true; break; END;
    if IsDebutParagraphe (TOBI,Niveau) then break;
  end;
end;

procedure SuprimeThisParagraphe(TobFacture, TOBL : TOB; var IndDep: integer);
var Niveau : integer;
    TOBI : TOB;
    StopItNow : boolean;
begin
  Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
  StopItNow := false;
  repeat
    TOBI := TOBFacture.detail[IndDep];
    if IsDebutParagraphe (TOBI,Niveau) then StopItNow := true;
    TOBI.free;
    Dec(IndDep);
  until (IndDep = 0 ) or (StopItNow);
end;

Procedure SupprimeOuvrageVide(TobIntermediaire : TOB);
Var Iind  : Integer;
    TOBI  : TOB;
begin

  Iind := TobIntermediaire.detail.Count-1;

  repeat
    TOBI := TobIntermediaire.detail[Iind];
    if (TOBI.detail.count = 0) and (TOBI.GetString('TYPE') = 'OUV') then FreeandNil(TOBI);
    Dec(Iind);
  until iInd < 0

end;

procedure StockeInfoTypeLigne (TOBL : TOB; ValLoc : T_Valeurs);
var QTe : double;
begin
  TOBL.PutValue('GLC_MTMOPA',ValLoc[23]);
  TOBL.PutValue('GLC_MTMOPR',ValLoc[24]);
  TOBL.PutValue('GLC_MTMOPV',ValLoc[25]);
  TOBL.PutValue('GLC_MTFOUPA',ValLoc[26]);
  TOBL.PutValue('GLC_MTFOUPR',ValLoc[27]);
  TOBL.PutValue('GLC_MTFOUPV',ValLoc[28]);
  TOBL.PutValue('GLC_MTINTPA',ValLoc[29]);
  TOBL.PutValue('GLC_MTINTPR',ValLoc[30]);
  TOBL.PutValue('GLC_MTINTPV',ValLoc[31]);
  TOBL.PutValue('GLC_MTLOCPA',ValLoc[32]);
  TOBL.PutValue('GLC_MTLOCPR',ValLoc[33]);
  TOBL.PutValue('GLC_MTLOCPV',ValLoc[34]);
  TOBL.PutValue('GLC_MTMATPA',ValLoc[35]);
  TOBL.PutValue('GLC_MTMATPR',ValLoc[36]);
  TOBL.PutValue('GLC_MTMATPV',ValLoc[37]);
  TOBL.PutValue('GLC_MTOUTPA',ValLoc[38]);
  TOBL.PutValue('GLC_MTOUTPR',ValLoc[39]);
  TOBL.PutValue('GLC_MTOUTPV',ValLoc[40]);
  TOBL.PutValue('GLC_MTSTPA',ValLoc[41]);
  TOBL.PutValue('GLC_MTSTPR',ValLoc[42]);
  TOBL.PutValue('GLC_MTSTPV',ValLoc[43]);
  TOBL.PutValue('GLC_MTAUTPA',ValLoc[44]);
  TOBL.PutValue('GLC_MTAUTPR',ValLoc[45]);
  TOBL.PutValue('GLC_MTAUTPV',ValLoc[46]);
end;

end.
