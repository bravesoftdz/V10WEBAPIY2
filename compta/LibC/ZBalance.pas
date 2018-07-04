unit ZBalance;

//=======================================================
//=================== Clés primaires ====================
//=======================================================
// HISTOBAL : HB_TYPEBAL,  HB_PLAN,  HB_EXERCICE, HB_DEVISE, HB_ETABLISSEMENT,
//            HB_DATE1, HB_DATE2, HB_COMPTE1,  HB_COMPTE2
//
//=======================================================
//==================== HB_TYPEBAL =======================
//=======================================================
// N1A | Balance N-1 avec A Nouveaux
// N1C | Balance N-1 pour comparatif
// N0A | Balance N (Reprise de balance)
// BDS | Balance de situation
//=======================================================
//====================== HB_PLAN ========================
//=======================================================
// BAL | Balance sans auxiliaire
// AUX | Balance avec auxiliaire
//=======================================================
//==================== E_NATUREGENE =====================
//=======================================================
// BQE | Banque
// CAI | Caisse
// CHA | Charge
// COC | Collectif client
// COD | Collectif divers
// COF | Collectif fournisseur
// COS | Collectif salarié
// DIV | Divers
// EXT | Extra-comptable
// IMO | Immobilisation
// PRO | Produit
// TIC | Tiers créditeurs
// TID | Tiers débiteurs
//=======================================================
//==================== T_NATUREAUXI =====================
//=======================================================
// AUC | Créditeur divers
// AUD | Débiteur divers
// CLI | Client
// DIV | Divers
// FOU | Fournisseur
// NCP | Non comptable
// SAL | Salarié
//=======================================================
//==================== J_NATUREJAL ======================
//=======================================================
// ACH | Achat
// ANA | ODA A-nouveaux
// ANO | A-nouveaux
// BQE | Banque
// CAI | Caisse
// CLO | Clôture
// ECC | Ecart de change
// EXT | Extra-comptable
// OD  | OD
// ODA | OD Analytique
// REG | Régularisation
// VTE | Vente

interface

uses Classes,
  {$IFDEF EAGLCLIENT}
     uTOB,
  {$ELSE}
     Db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
     SysUtils,
     ComCtrls,
     Forms,
     Controls,
     HCtrls,
     HEnt1,
     HStatus,
     ParamSoc,
     Ent1,
     ed_tools,
     SaisUtil,
     SaisComm,
     HCompte,
     TZ,
     ZTypes,
  {$IFDEF MODENT1}
     CPTypeCons,
  {$ENDIF MODENT1}
     hmsgBox; // PgiError

procedure FillComboBds(Combo : TControl) ;

type
  TZBalance = class
  private
    Ecrs   : TZF ;
    // Références du where pour l'objet
    FDev : RDevise ;
    FOpt : ROpt ;
    FNumPiece : Integer ;
    DateDeb, DateFin : TDateTime ; bAux : Boolean ;
    function    SoldeD(ColD, ColC : string ; Ecr : TZF; Q : TQuery) : Double ;
    function    SoldeC(ColD, ColC : string ; Ecr : TZF; Q : TQuery) : Double ;
    function    GetCount : Integer ;
    function    GetIsEcr : Boolean ;
  public
    constructor Create(Dev : RDevise; Opt : ROpt) ;
    destructor  Destroy ; override ;
    procedure   SetOptions(Opt : ROpt) ;
    function    CreateRow(Q : TQuery; Row : LongInt) : TZF ;
    function    CreateRowAux(Q : TQuery; NumCompte : string) : TZF ;
    procedure   DeleteRow(Row : LongInt) ;
    function    GetRow(Row : LongInt) : TZF ;
    function    FindCompte(var NumCompte, NumAux : string) : Integer ;
    function    GetDateDeb : TDateTime ;
    function    GetDateFin : TDateTime ;
    function    GetAux : Boolean ;
    function    Exist(bMaj : Boolean) : Boolean ;
    function    ExistBds(bMaj : Boolean) : Boolean ;
    function    ExistEcrs : Boolean ;
    function    EcrsModif : Integer ;
    function    Read(CptePerte: string=''; CpteBenef: string='') : Boolean ;
    function    Write(CptePerte: string=''; CpteBenef: string='') : Boolean ;
    function    Del : Boolean ;
    // A Nouveaux
    function    AnDel : Boolean ;
    function    AnoPrepare(EcrMaster : TZF; JalAno, sMode: string; DateEcr: TDateTime;
                           NumPiece: Integer; var k: Integer; bLettrage: Boolean;
                           SOpp : Double=0; bEcart :Boolean=FALSE) : TZF ;
    function    AnWrite(bLettrage: Boolean) : Boolean ; //(Resultat : Double; CpteResultat : string) : Boolean ;
    function    AnOk : Boolean ;
    {$IFNDEF EAGLCLIENT}
    function    AnFabricReqINV(fb : TFichierBase) : TQuery ;
    {$ENDIF}
    // Balance de reprise
    function    RbDel : Boolean ;
    function    RbRead : Boolean ;
    function    RbWrite : Boolean ;
    function    RbPrepare(EcrMaster : TZF; JalRb, sMode: string; DateEcr: TDateTime;
                          NumPiece: Integer; var k: Integer;  SOpp : Double=0; bEcart :Boolean=FALSE) : TZF ;
    // Génération de balance
    function    GenAuto : Boolean ;
    function    GenAuto2 : Boolean ;
    property    Count : Integer read GetCount ;
    property    Ecr   : Boolean read GetIsEcr ;
  end ;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  ZPlan;

//=======================================================
//================== Interface commune ==================
//=======================================================
procedure FillComboBds(Combo : TControl) ;
var Q : TQuery ; Cols, Where : string ;
begin
if not (Combo is THValComboBox) then Exit ;
THValComboBox(Combo).Clear ; THValComboBox(Combo).Values.Clear ;
Cols:='HB_EXERCICE, HB_DATE1, HB_DATE2' ;
Where:='HB_TYPEBAL="BDS"' ;
Q:=OpenSQL('SELECT DISTINCT '+Cols+' FROM HISTOBAL WHERE '+Where, TRUE) ;
while not Q.EOF do
  begin
  THValComboBox(Combo).Values.Add(Q.Fields[0].AsString+';'+Q.Fields[1].AsString+';'+Q.Fields[2].AsString+';') ;
  THValComboBox(Combo).Items.Add(TraduireMemoire('Balance du '+DateToStr(Q.Fields[1].AsDateTime)+' au '+DateToStr(FinDeMois(Q.Fields[2].AsDateTime)))) ;
  Q.Next ;
  end ;
Ferme(Q) ;
THValComboBox(Combo).ItemIndex:=0 ;
end ;

//=======================================================
//=================== Objet TZBalance ===================
//=======================================================
constructor TZBalance.Create(Dev : RDevise; Opt : ROpt) ;
begin
if Opt.bEcr then Ecrs:=TZF.Create('VRB',      nil, -1)
            else Ecrs:=TZF.Create('VBALANCE', nil, -1) ;
FDev:=Dev ;
FOpt:=Opt ;
FNumPiece:=-1 ;
end ;

destructor TZBalance.Destroy ;
begin
if Ecrs<>nil then Ecrs.Free ;
end ;

procedure TZBalance.SetOptions(Opt : ROpt) ;
begin
FOpt:=Opt ;
end ;

function TZBalance.GetCount : Integer ;
begin
Result:=Ecrs.Detail.Count ;
end ;

function TZBalance.GetIsEcr : Boolean ;
begin
Result:=FOpt.bEcr ;
end ;

function TZBalance.CreateRow(Q : TQuery; Row : LongInt) : TZF ;
var Ecr : TZF ; RowRef : LongInt ; TableD : string ;
begin
if Row>Ecrs.Detail.Count-1 then RowRef:=-1
                           else RowRef:=Row ;
if FOpt.bEcr then TableD:='ECRITURE' else TableD:='HISTOBAL' ;
if Q<>nil then Ecr:=TZF.CreateDB(TableD, Ecrs, RowRef, Q)
          else Ecr:=TZF.Create(TableD, Ecrs, RowRef) ;
Result:=Ecr ;
end ;

function TZBalance.CreateRowAux(Q : TQuery; NumCompte : string) : TZF ;
var Ecr : TZF ; iFind : Integer ; NumAux, TableD : string ;
begin
NumAux:='' ; TableD:='' ;
iFind:=FindCompte(NumCompte, NumAux) ;
if iFind<0 then begin Result:=nil ; Exit ; end ;
if FOpt.bEcr then TableD:='ECRITURE' else TableD:='HISTOBAL' ;
if Q<>nil then Ecr:=TZF.CreateDB(TableD, GetRow(iFind), -1, Q)
          else Ecr:=TZF.Create(TableD, GetRow(iFind), -1) ;
Result:=Ecr ;
end ;

procedure TZBalance.DeleteRow(Row : LongInt) ;
var Ecr : TZF ;
begin
if Row>Ecrs.Detail.Count-1 then Exit ;
Ecr:=TZF(Ecrs.Detail[Row]) ; Ecr.Free ;
end ;

function TZBalance.GetRow(Row : LongInt) : TZF ;
begin
Result:=nil ; if Row>Ecrs.Detail.Count-1 then Exit ; Result:=TZF(Ecrs.Detail[Row]) ;
end ;

function TZBalance.FindCompte(var NumCompte, NumAux : string) : Integer ;
var i : integer ; Col1Find, Col2Find : string ;
begin
Result:=-1 ;
NumCompte:=Trim(NumCompte) ; if NumCompte='' then Exit ;
NumCompte:=BourreLaDonc(NumCompte, fbGene) ;
if NumAux<>'' then NumAux:=BourreLaDonc(NumAux, fbAux) ;
if FOpt.bEcr then begin Col1Find:='E_GENERAL' ;  Col2Find:='E_AUXILIAIRE' ; end
             else begin Col1Find:='HB_COMPTE1' ; Col2Find:='HB_COMPTE2' ;   end ;
for i:=0 to Ecrs.Detail.Count-1 do
  if (TZF(Ecrs.Detail[i]).GetValue(Col1Find)=NumCompte) and
     (TZF(Ecrs.Detail[i]).GetValue(Col2Find)=NumAux) then
     begin Result:=i ; Exit ; end ;
end ;

function TZBalance.GetDateDeb : TDateTime ;
begin Result:=DateDeb ; end ;

function TZBalance.GetDateFin : TDateTime ;
begin Result:=DateFin ; end ;

function TZBalance.GetAux : Boolean ;
begin Result:=bAux ; end ;

function TZBalance.Exist(bMaj : Boolean) : Boolean ;
var Q : TQuery ; Cols : string ;
begin
if bMaj then Cols:='HB_EXERCICE, HB_DATE1, HB_DATE2, HB_PLAN'
        else Cols:='HB_EXERCICE' ;
Q:=OpenSQL('SELECT '+Cols+' FROM HISTOBAL'
          +' WHERE HB_TYPEBAL="'+FOpt.TypeBal+'"'
          +' AND HB_EXERCICE="'+FOpt.ExoBal+'"', TRUE) ;
Result:=not Q.EOF ;
if (Result) and (bMaj) then
   begin
   DateDeb:=Q.Fields[1].AsDateTime ;
   DateFin:=Q.Fields[2].AsDateTime ;
   bAux:=(Q.Fields[3].AsString='AUX') ;
   end ;
Ferme(Q) ;
end ;

function TZBalance.ExistBds(bMaj : Boolean) : Boolean ;
var Q : TQuery ; Cols : string ;
begin
if bMaj then Cols:='HB_EXERCICE, HB_DATE1, HB_DATE2, HB_PLAN'
        else Cols:='HB_EXERCICE' ;
Q:=OpenSQL('SELECT DISTINCT '+Cols+' FROM HISTOBAL WHERE HB_TYPEBAL="'+FOpt.TypeBal+'"'
          +' AND HB_EXERCICE="'+FOpt.ExoBal+'"'
          +' AND HB_DATE1="'+USDateTime(EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour))+'"'
          +' AND HB_DATE2="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"',
          TRUE) ;
Result:=not Q.EOF ;
if (Result) and (bMaj) then
   begin
   DateDeb:=Q.Fields[1].AsDateTime ;
   DateFin:=Q.Fields[2].AsDateTime ;
   bAux:=(Q.Fields[3].AsString='AUX') ;
   end ;
Ferme(Q) ;
end ;

function TZBalance.ExistEcrs : Boolean ;
var Q : TQuery ; Where, Cols : string ;
begin
Cols:='E_EXERCICE' ;
Where:=' E_EXERCICE="'+FOpt.ExoBal+'"'
      +' AND E_QUALIFORIGINE<>"'+FOpt.TypeBal+'"'
      +' AND E_CREERPAR<>"DET"' ;
Q:=OpenSQL('SELECT DISTINCT '+Cols+' FROM ECRITURE WHERE'+Where, TRUE) ;
Result:=not Q.EOF ;
Ferme(Q) ;
end ;

// Cherche une écriture générée par la balance N et modifiée par la saisie
function TZBalance.EcrsModif : Integer ;
var Q : TQuery ; Where, Cols : string ;
begin
Result:=-1 ;
Cols:='E_NUMEROPIECE, E_DATECREATION, E_DATEMODIF' ;
Where:=' E_EXERCICE="'+FOpt.ExoBal+'"'
      +' AND E_QUALIFORIGINE="'+FOpt.TypeBal+'"'
      +' AND E_CREERPAR="BAL" AND (E_DATECREATION<>E_DATEMODIF OR E_VALIDE="X")' ;
Q:=OpenSQL('SELECT '+Cols+' FROM ECRITURE WHERE'+Where, TRUE) ;
if not Q.EOF then Result:=Q.Fields[0].AsInteger ;
Ferme(Q) ;
end ;

function TZBalance.Read(CptePerte: string=''; CpteBenef: string='') : Boolean ;
var QEcr : TQuery ; Where : string ;
begin
if FOpt.bEcr then begin Result:=RbRead ; Exit ; end ;
Where:='HB_TYPEBAL="'+FOpt.TypeBal+'"'
      +' AND HB_EXERCICE="'+FOpt.ExoBal+'"'
      +' AND HB_DATE1="'+USDateTime(EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour))+'"'
      +' AND HB_DATE2="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"' ;
if CptePerte<>'' then Where:=Where+' AND HB_COMPTE1<>"'+CptePerte+'"' ;
if CpteBenef<>'' then Where:=Where+' AND HB_COMPTE1<>"'+CpteBenef+'"' ;
QEcr:=OpenSQL('SELECT * FROM HISTOBAL WHERE '+Where
             +' ORDER BY HB_TYPEBAL, HB_EXERCICE, HB_COMPTE1', TRUE) ;
Result:=not QEcr.EOF ;
while not QEcr.EOF do
  begin
  // Charger le TZF
  CreateRow(QEcr, Ecrs.Detail.Count+1) ;
  QEcr.Next ;
  end ;
Ferme(QEcr) ;
end ;

function TZBalance.Write(CptePerte: string; CpteBenef: string) : Boolean ;
var i, k : Integer ;
begin
Result:=TRUE ;
InitMove(Ecrs.Detail.Count, TraduireMemoire('Ecriture de la balance')) ;
for i:=0 to Ecrs.Detail.Count-1 do
  begin
  MoveCur(FALSE) ;
  if (CptePerte<>'') and (TZF(Ecrs.Detail[i]).GetValue('HB_COMPTE1')=CptePerte) then Continue ;
  if (CpteBenef<>'') and (TZF(Ecrs.Detail[i]).GetValue('HB_COMPTE1')=CpteBenef) then Continue ;
  if TZF(Ecrs.Detail[i]).Detail.Count>0 then
    begin
    for k:=0 to TZF(Ecrs.Detail[i]).Detail.Count-1 do
      begin
      TZF(TZF(Ecrs.Detail[i]).Detail[k]).InsertDB(nil) ;
      end ;
    Continue ;
    end ;
  if TZF(Ecrs.Detail[i]).GetValue('BADROW')<>'X' then
    begin
    TZF(Ecrs.Detail[i]).InsertDB(nil) ; 
    end ;
  end ;
FiniMove ;
end ;

function TZBalance.Del : Boolean ;
var Where : string ;
begin
//try
Result:=TRUE ;
Where:=' HB_TYPEBAL="'+FOpt.TypeBal+'"'
      +' AND HB_EXERCICE="'+FOpt.ExoBal+'"'
      +' AND HB_DATE1="'+USDateTime(EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour))+'"'
      +' AND HB_DATE2="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"' ;
if ExecuteSQL('DELETE FROM HISTOBAL WHERE'+Where)<=0 then
   begin Result:=FALSE ; V_PGI.IOError:=oeSaisie ; Exit ; end ;
//except
//  on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : TZBalance.Del');
//end;

end ;

//=======================================================
//================== Génération des AN ==================
//=======================================================
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/02/2004
Modifié le ... :   /  /
Description .. : Passage en E-AGL
Suite ........ : E_QUALIFORIGINE  : COMBO => FOpt.TypeBal
Mots clefs ... : 
*****************************************************************}
function TZBalance.AnDel : Boolean ;
var QEcr, QAna : TQuery ; Where, JalAno, Cols, WhereAna, ColsAna : string ;
    JalSoldeD, JalSoldeC : Double ; bAna : Boolean ;
    lDD, lCD : Double;
    lFRM : TFRM;
    ll : Longint;
begin
  Result := TRUE ;
  bAna := False;
  JalAno := '' ;
  Where  := ' E_QUALIFORIGINE = "' + FOpt.TypeBal + '"' +
            ' AND E_EXERCICE = "' + FOpt.ExoAno + '"' +
            ' AND E_DATECOMPTABLE >= "' + USDateTime(EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour)) + '"' +
            ' AND E_DATECOMPTABLE <= "' + USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour)+FOpt.DecalAno) + '"';

  // Liste des champs
  Cols := 'E_GENERAL, E_AUXILIAIRE, E_DEBIT, E_CREDIT, E_JOURNAL, E_NUMEROPIECE' ;

  QEcr := OpenSQL('SELECT ' + Cols + ' FROM ECRITURE WHERE' + Where, True) ;
  if not QEcr.EOF then
  begin
    JalAno    := QEcr.FindField('E_JOURNAL').AsString ;
    FNumPiece := QEcr.FindField('E_NUMEROPIECE').AsInteger ;
  end
  else
    FNumPiece := -1 ;

  // Comptes généraux
  JalSoldeD := 0 ;
  JalSoldeC := 0 ;
  while not QEcr.EOF do
  begin
    lDD := SoldeD('E_DEBIT', 'E_CREDIT', nil, QEcr);
    lCD := SoldeC('E_DEBIT', 'E_CREDIT', nil, QEcr);

    JalSoldeD := JalSoldeD + lDD;
    JalSoldeC := JalSoldeC + lCD;

    if ExecuteSQL('UPDATE GENERAUX SET ' +
                      'G_TOTDEBANO   = ' + VariantToSql(lDD) + ',' +
                      'G_TOTCREANO   = ' + VariantToSql(lCD) + ',' +
                      'G_TOTALDEBIT  = ' + VariantToSql(lDD) + ',' +
                      'G_TOTALCREDIT = ' + VariantToSql(lCD) + ' ' +
                      'WHERE G_GENERAL = "' + QEcr.FindField('E_GENERAL').AsString + '"') <> 1 then
    begin
      Result := False ;
      V_PGI.IoError := oeSaisie ;
    end;
    QEcr.Next;
  end;

  // Comptes tiers
  QEcr.First ;
  while not QEcr.EOF do
  begin
    if QEcr.FindField('E_AUXILIAIRE').AsString = '' then
    begin
      QEcr.Next ;
      Continue ;
    end ;

    lDD := SoldeD('E_DEBIT', 'E_CREDIT', nil, QEcr) ;
    lCD := SoldeC('E_DEBIT', 'E_CREDIT', nil, QEcr) ;

    if ExecuteSql('UPDATE TIERS SET ' +
                  'T_TOTDEBANO   = ' + VariantToSql(lDD) + ',' +
                  'T_TOTCREANO   = ' + VariantToSql(lCD) + ',' +
                  'T_TOTALDEBIT  = ' + VariantToSql(lDD) + ',' +
                  'T_TOTALCREDIT = ' + VariantToSql(lCD) + ' ' +
                  'WHERE T_AUXILIAIRE = "' + QEcr.FindField('E_AUXILIAIRE').AsString + '"') <> 1 then
    begin
      Result := False ;
      V_PGI.IoError := oeSaisie ;
    end;
    QEcr.Next;
  end;

  // Sections
  WhereAna := ' Y_NUMEROPIECE = ' + IntToStr(FNumPiece) +
              ' AND Y_JOURNAL = "' + JalAno + '"' +
              ' AND Y_EXERCICE= "' + FOpt.ExoAno + '"' +
              ' AND Y_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour)) + '"' +
              ' AND Y_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour)+FOpt.DecalAno) + '"';

  ColsAna := 'Y_SECTION, Y_AXE, Y_DEBIT, Y_CREDIT' ;
  try
    try
      QAna := OpenSQL('SELECT ' + ColsAna + ' FROM ANALYTIQ WHERE' + WhereAna, True) ;
      bAna := not QAna.EOF ;
      while not QAna.EOF do
      begin
        Fillchar(lFRM, SizeOf(lFRM),#0) ;
        lFRM.Cpt := QAna.FindField('Y_SECTION').AsString ;
        lFRM.Axe := QAna.FindField('Y_AXE').AsString ;
        lFRM.Deb := QAna.FindField('Y_DEBIT').AsFloat;
        lFRM.Cre := QAna.FindField('Y_CREDIT').AsFloat;

        if ((lFRM.Axe <> 'G') and (lFRM.Axe <> 'T')) then
        begin
          AttribParamsNew(lFRM, QAna.FindField('Y_DEBIT').AsFloat, QAna.FindField('Y_CREDIT').AsFloat, FOpt.TypeExo);
          ll := ExecReqINVNew(fbSect, lFRM) ;
          if ll <> 1 then
            V_PGI.IoError := oeSaisie ;
        end;
        QAna.Next ;
      end;
    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : AnDel -> Section');
    end;
  finally
    Ferme(QAna) ;
  end ;

  // Le journal
  QEcr.First ;
  if not QEcr.EOF then
  begin
    if ExecuteSql('UPDATE JOURNAL SET ' +
                  'J_DEBITDERNMVT  = ' + VariantToSql(0) + ', ' +
                  'J_CREDITDERNMVT = ' + VariantToSql(0) + ', ' +
                  'J_DATEDERNMVT   = "' + UsDateTime((EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour) + FOpt.DecalAno)) + '", ' +
                  'J_NUMDERNMVT    = ' + VariantToSql(FNumPiece) + ', ' +
                  'J_TOTALDEBIT    = ' + VariantToSql(0) + ', ' +
                  'J_TOTALCREDIT   = ' + VariantToSql(0) + ', ' +
                  'J_TOTDEBE       = ' + VariantToSql(JalSoldeD) + ', ' +
                  'J_TOTCREE       = ' + VariantToSql(JalSoldeC) + ', ' +
                  'J_TOTDEBS       = ' + VariantToSql(0) + ', ' +
                  'J_TOTCRES       = ' + VariantToSql(0) + ', ' +
                  'J_TOTDEBP       = ' + VariantToSql(0) + ', ' +
                  'J_TOTCREP       = ' + VariantToSql(0) + ' ' +
                  'WHERE J_JOURNAL = "' + JalAno + '"') <> 1 then
    begin
      Result := False ;
      V_PGI.IoError := oeSaisie ;
    end;
    QEcr.Next;
  end;
  Ferme(QEcr) ;

  // Destruction des écritures
  if ExecuteSQL('DELETE FROM ECRITURE WHERE'+Where) <= 0 then
  begin
    Result := FALSE ;
    V_PGI.IOError := oeSaisie ;
    Exit ;
  end ;

  // Destruction des écritures analytiques
  if bAna then
    if ExecuteSQL('DELETE FROM ANALYTIQ WHERE'+WhereAna) <= 0 then
    begin
      Result := FALSE ;
      V_PGI.IOError := oeSaisie ;
      Exit ;
    end ;
end ;

function TZBalance.SoldeD(ColD, ColC : string ; Ecr : TZF; Q : TQuery) : Double ;
var Solde : Double ;
begin
Result:=0 ;
if Q=nil then Solde:=Ecr.GetValue(ColD)-Ecr.GetValue(ColC)
         else Solde:=Q.FindField(ColD).AsFloat-Q.FindField(ColC).AsFloat ;
if Solde>0 then Result:=Solde ;
end ;

function TZBalance.SoldeC(ColD, ColC : string ; Ecr : TZF; Q : TQuery) : Double ;
var Solde : Double ;
begin
Result:=0 ;
if Q=nil then Solde:=Ecr.GetValue(ColC)-Ecr.GetValue(ColD)
         else Solde:=Q.FindField(ColC).AsFloat-Q.FindField(ColD).AsFloat ;
if Solde>0 then Result:=Solde ;
end ;

function TZBalance.AnoPrepare(EcrMaster : TZF; JalAno, sMode: string; DateEcr: TDateTime;
                              NumPiece: Integer; var k: Integer; bLettrage: Boolean;
                              SOpp : Double=0; bEcart :Boolean=FALSE) : TZF ;
var Q : TQuery ; EcrAno : TZF ; z : Integer ; Solde : Double ;
    Mode, DecEnc : string ; bLe, bPo : Boolean ;
begin
EcrAno:=TZF.Create('ECRITURE', nil, -1) ;
EcrAno.SetDateModif(Date) ;
for z:=1 to MAXAXE do TZF.Create('A'+IntToStr(z), EcrAno, -1) ;
EcrAno.PutValue('E_EXERCICE',      FOpt.ExoAno) ;
EcrAno.PutValue('E_JOURNAL',       JalAno) ;
EcrAno.PutValue('E_DATECOMPTABLE', DateEcr) ;
EcrAno.PutValue('E_NUMEROPIECE',   NumPiece) ;
EcrAno.PutValue('E_NUMLIGNE',      k); k:=k+1 ;
EcrAno.PutValue('E_GENERAL',       EcrMaster.GetValue('HB_COMPTE1')) ;
EcrAno.PutValue('E_AUXILIAIRE',    EcrMaster.GetValue('HB_COMPTE2')) ;
EcrAno.PutValue('E_DEBIT',         SoldeD('HB_DEBIT', 'HB_CREDIT', EcrMaster, nil)) ;
EcrAno.PutValue('E_CREDIT',        SoldeC('HB_DEBIT', 'HB_CREDIT', EcrMaster, nil)) ;
EcrAno.PutValue('E_DEBITDEV',      SoldeD('HB_DEBITDEV', 'HB_CREDITDEV', EcrMaster, nil)) ;
EcrAno.PutValue('E_CREDITDEV',     SoldeC('HB_DEBITDEV', 'HB_CREDITDEV', EcrMaster, nil)) ;
EcrAno.PutValue('E_NATUREPIECE',   'OD') ;
EcrAno.PutValue('E_TYPEMVT',       'DIV') ;
EcrAno.PutValue('E_QUALIFPIECE',   'N') ;
EcrAno.PutValue('E_VALIDE',        'X') ;
EcrAno.PutValue('E_DEVISE',        EcrMaster.GetValue('HB_DEVISE')) ;
EcrAno.PutValue('E_TAUXDEV',       1) ;
EcrAno.PutValue('E_QUALIFORIGINE', FOpt.TypeBal) ;
if bLettrage then EcrAno.PutValue('E_ECRANOUVEAU', 'H')
             else EcrAno.PutValue('E_ECRANOUVEAU', 'OAN') ;
EcrAno.PutValue('E_MODESAISIE',    sMode) ;
EcrAno.PutValue('E_PERIODE',       GetPeriode(DateEcr)) ;
EcrAno.PutValue('E_SEMAINE',       NumSemaine(DateEcr)) ;
bLe:=(Copy(EcrMaster.GetValue('LETTRABLE'), 1, 1)='X') ;
bPo:=(Copy(EcrMaster.GetValue('LETTRABLE'), 2, 1)='X') ;
if (bLe) or (bPo) then
  begin
  Mode:='' ;
  Q:=OpenSQL('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="'
             +EcrMaster.GetValue('MODEREGL')+'"', TRUE) ;
  if not Q.EOF then Mode:=Q.FindField('MR_MP1').AsString ;
  Ferme(Q) ;
  if Mode='' then
    begin
    Q:=OpenSQL('SELECT MP_MODEPAIE FROM MODEPAIE', TRUE) ;
    if not Q.EOF then Mode:=Q.FindField('MP_MODEPAIE').AsString ;
    Ferme(Q) ;
    end ;
  EcrAno.PutValue('E_MODEPAIE',        Mode) ;
  EcrAno.PutValue('E_DATEECHEANCE',    DateEcr) ;
  EcrAno.PutValue('E_ORIGINEPAIEMENT', DateEcr) ;
  EcrAno.PutValue('E_ECHE',            'X') ;
  EcrAno.PutValue('E_NUMECHE',         1) ;
  if bLe then
    begin
    EcrAno.PutValue('E_ETATLETTRAGE',    'AL') ;
    Solde:=EcrAno.GetValue('E_DEBIT')-EcrAno.GetValue('E_CREDIT') ;
    if Solde=0 then DecEnc:='' ;
    if Solde>0 then // Débiteurs
      begin
      if EcrMaster.GetValue('TYPEGENE')='COC' then DecEnc:='ENC' else
        if EcrMaster.GetValue('TYPEGENE')='COF' then DecEnc:='DEC' else
          DecEnc:='ENC' ;
      end else // Créditeurs
      begin
      if EcrMaster.GetValue('TYPEGENE')='COC' then DecEnc:='ENC' else
        if EcrMaster.GetValue('TYPEGENE')='COF' then DecEnc:='DEC' else
          DecEnc:='DEC' ;
      end ;
    EcrAno.PutValue('E_ENCAISSEMENT', DecEnc) ;
    EcrAno.PutValue('E_RIB', GetRIBPrincipal(EcrMaster.GetValue('HB_COMPTE2'))) ;
    end ;
  EcrAno.PutValue('E_DATEVALEUR', DateEcr) ;
  end ;
if VentilerTOB(EcrAno, '', 0, FDev.Decimale, FALSE) then EcrAno.PutValue('E_ANA', 'X') ;
EcrAno.InsertDB(nil) ;
Result:=EcrAno ;
end ;

function TZBalance.AnWrite(bLettrage: Boolean) : Boolean ; //(Resultat : Double; CpteResultat : string) : Boolean ;
var Q :TQuery ; EcrAno, EcrMaster, ee : TZF ; JalSoldeD, JalSoldeC : Double ;
    DateEcr : TDateTime ; i, j, k, c : Integer ; e : TList ;
    JalAno, sCompteur, sMode : string ;
    lFRM : TFRM;
    ll : LongInt;
begin
Result:=TRUE ;
DateEcr:=EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour)+FOpt.DecalAno ;
// Journal des AN
{$IFDEF SPEC302}
Q:=OpenSQL('SELECT SO_JALOUVRE FROM SOCIETE', TRUE) ;
JalAno:=Q.Fields[0].AsString ;
Ferme(Q) ;
{$ELSE}
JalAno:=GetParamSoc('SO_JALOUVRE') ;
{$ENDIF}
// Numéro de pièce à uiliser
//Critere:='E_JOURNAL="'+JalAno+'" AND E_NUMLIGNE=1' ;
//Q:=OpenSQL('SELECT DISTINCT E_NUMEROPIECE FROM ECRITURE WHERE '+Critere
//          +' ORDER BY E_NUMEROPIECE DESC', TRUE) ;
//NumPiece:=Q.Fields[0].AsInteger+1 ;
//Ferme(Q) ;
if FNumPiece<0 then
  begin
  Q:=OpenSQL('SELECT J_COMPTEURNORMAL, J_MODESAISIE FROM JOURNAL WHERE J_JOURNAL="'+JalAno+'"', TRUE) ;
  sCompteur:=Q.Fields[0].AsString ;
  sMode:=Q.Fields[1].AsString ;
  Ferme(Q) ;
  SetIncNum(EcrGen, sCompteur, FNumPiece, DateEcr) ;
  end ;
// Let's go !!!
InitMove(Ecrs.Detail.Count, TraduireMemoire('Génération des A Nouveaux')) ;
k:=1 ; e:=TList.Create ;
for i:=0 to Ecrs.Detail.Count-1 do
    begin
    MoveCur(FALSE) ;
    EcrMaster:=TZF(Ecrs.Detail[i]) ;
    (*
    if (EcrMaster.GetValue('HB_COMPTE1')=CpteResultat) and (Resultat<>0) then
      begin
      EcrMaster.PutValue('BADROW', '-') ;
      EcrMaster.PutValue('BADROW', '-') ;
      end ;
    *)
    if EcrMaster.GetValue('BADROW'  )='X'   then Continue ;
    if EcrMaster.GetValue('TYPEGENE')='CHA' then Continue ;
    if EcrMaster.GetValue('TYPEGENE')='PRO' then Continue ;
    if EcrMaster.GetValue('TYPEGENE')='EXT' then Continue ;
    if EcrMaster.Detail.Count>0 then
      begin
      for j:=0 to EcrMaster.Detail.Count-1 do
        begin
        EcrAno:=AnoPrepare(TZF(EcrMaster.Detail[j]), JalAno, sMode, DateEcr, FNumPiece, k, bLettrage) ;
        e.Add(EcrAno) ;
        end ;
      end else
      begin
      EcrAno:=AnoPrepare(EcrMaster, JalAno, sMode, DateEcr, FNumPiece, k, bLettrage) ;
      e.Add(EcrAno) ;
      end ;
    end ;
FiniMove ;

  // Les comptes généraux
  JalSoldeD:=0 ; JalSoldeC:=0 ;
  InitMove(e.Count, TraduireMemoire('Mise à jour des comptes généraux')) ;

  for i:=0 to e.Count-1 do
  begin
    MoveCur(FALSE) ;
    if TZF(e[i]).GetValue('BADROW')='X' then Continue ;
    Fillchar(lFRM, SizeOf(lFRM),#0) ;
    lFRM.Cpt := TZF(e[i]).GetValue('E_GENERAL');
    lFRM.Deb := TZF(e[i]).GetValue('E_DEBIT')  ;
    lFRM.Cre := TZF(e[i]).GetValue('E_CREDIT') ;

    JalSoldeD := JalSoldeD + TZF(e[i]).GetValue('E_DEBIT');
    JalSoldeC := JalSoldeC + TZF(e[i]).GetValue('E_CREDIT');

    ll := ExecReqMAJ( fbGene, True, False, lFRM ) ;
    if ll <> 1 then
    begin
      Result := False;
      V_PGI.IoError := oeSaisie ;
    end;

  end ;
  FiniMove ;

  // Les comptes tiers
  InitMove(e.Count, TraduireMemoire('Mise à jour des comptes tiers')) ;
  for i:=0 to e.Count-1 do
  begin
    MoveCur(FALSE) ;
    if TZF(e[i]).GetValue('BADROW')='X' then Continue ;
    if TZF(e[i]).GetValue('E_AUXILIAIRE')='' then Continue ;
    Fillchar(lFRM, SizeOf(lFRM),#0) ;
    lFRM.Cpt := TZF(e[i]).GetValue('E_AUXILIAIRE') ;
    lFRM.Deb := TZF(e[i]).GetValue('E_DEBIT')  ;
    lFRM.Cre := TZF(e[i]).GetValue('E_CREDIT') ;

    ll := ExecReqMAJ( fbAux, True, False, lFRM ) ;
    if ll <> 1 then
    begin
      Result := False;
      V_PGI.IoError := oeSaisie ;
    end;
  end ;
  FiniMove ;

  // Sections
  InitMove(e.Count, TraduireMemoire('Mise à jour des sections')) ;
  for i:=0 to e.Count-1 do
  begin
    MoveCur(FALSE) ;
    if TZF(e[i]).GetValue('BADROW') = 'X' then Continue ;
    if TZF(e[i]).GetValue('E_ANA') <> 'X' then Continue ;
    for j := 0 to MAXAXE-1 do
    begin
      c := TZF(e[i]).Detail[j].Detail.Count ;
      if c = 0 then Continue ;
      for k := 0 to c-1 do
      begin
        ee := TZF(TZF((TZF(e[i]).Detail[j]).Detail[k])) ;
        Fillchar(lFRM, SizeOf(lFRM),#0) ;
        lFRM.Cpt := ee.GetValue('Y_SECTION') ;
        lFRM.Axe := ee.GetValue('Y_AXE') ;
        lFRM.Deb := ee.GetValue('Y_DEBIT') ;
        lFRM.Cre := ee.GetValue('Y_CREDIT') ;
        ll := ExecReqMAJ( fbSect, True, False, lFRM ) ;
        if ll <> 1 then
        begin
          Result := False;
          V_PGI.IoError := oeSaisie ;
        end;
      end;
    end ;
  end ;
  FiniMove ;

  // Le journal
  Fillchar(lFRM, SizeOf(lFRM),#0) ;
  lFRM.Cpt   := JalAno ;
  lFRM.NumD  := FNumPiece ;
  lFRM.DateD := EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour) + FOpt.DecalAno ;
  lFRM.DE    := JalSoldeD ;
  lFRM.CE    := JalSoldeC ;
  lFRM.Deb   := 0;
  lFRM.Cre   := 0;
  lFRM.DS    := 0;
  lFRM.CS    := 0;
  lFRM.DP    := 0;
  lFRM.CP    := 0;
  ll := ExecReqMaj( fbJal, True, False, lFRM);
  if ll <> 1 then
  begin
    Result := False;
    V_PGI.IoError := oeSaisie ;
  end;

  VideListe(e) ;
  e.Free ;

  if bLettrage then
  begin
    if ExecuteSQL('UPDATE SOCIETE SET SO_EXOV8="'+FOpt.ExoAno+'"')<=0 then
    begin
      Result:=FALSE ;
      V_PGI.IOError:=oeSaisie ; Exit ;
    end
    else
    begin
      VH^.ExoV8.Code:=FOpt.ExoAno ;
      VH^.ExoV8.Deb:=EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour) ;
      VH^.ExoV8.Fin:=EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour)+FOpt.DecalAno ;
    end ;
  end
  else
  begin
    if ExecuteSQL('UPDATE SOCIETE SET SO_EXOV8=""')<=0 then
    begin
      Result := FALSE ;
      V_PGI.IOError := oeSaisie ;
      Exit ;
    end
    else
    begin
      VH^.ExoV8.Code := '' ;
      VH^.ExoV8.Deb  := EncodeDate(1900, 1, 1) ;
      VH^.ExoV8.Fin  := EncodeDate(1900, 1, 1) ;
    end ;
  end ;
  FNumPiece := -1 ;
end ;

function TZBalance.AnOk : Boolean ;
var Q : TQuery ; Where : string ;
begin
Where:=' E_QUALIFORIGINE="'+FOpt.TypeBal+'"'
      +' AND E_EXERCICE="'+FOpt.ExoBal+'"'
      +' AND E_LETTRAGE=""'
      +' AND E_REFPOINTAGE=""';
Q:=OpenSQL('SELECT E_QUALIFORIGINE FROM ECRITURE WHERE'+Where, TRUE) ;
Result:=Q.EOF ;
Ferme(Q) ;
end ;

{$IFNDEF EAGLCLIENT}
function TZBalance.AnFabricReqINV(fb : TFichierBase) : TQuery ;
var Q : TQuery ; SQL : string ;
begin
  Q:=TQuery.Create(Application) ; Q.DataBaseName:=DBSOC.DataBaseName ;
case fb of
   fbGene : SQL:='UPDATE GENERAUX SET G_TOTDEBANO=G_TOTDEBANO-:DD, G_TOTCREANO=G_TOTCREANO-:CD, '
                +'G_TOTALDEBIT=G_TOTALDEBIT-:DD, G_TOTALCREDIT=G_TOTALCREDIT-:CD WHERE G_GENERAL=:CPTE' ;

    fbAux : SQL:='UPDATE TIERS SET T_TOTDEBANO=T_TOTDEBANO-:DD, T_TOTCREANO=T_TOTCREANO-:CD, '
                +'T_TOTALDEBIT=T_TOTALDEBIT-:DD, T_TOTALCREDIT=T_TOTALCREDIT-:CD WHERE T_AUXILIAIRE=:CPTE' ;

    fbJal : SQL:='UPDATE JOURNAL SET J_DEBITDERNMVT=:DD,  J_CREDITDERNMVT=:CD, '
                +'J_DATEDERNMVT=:DATED, J_NUMDERNMVT=:NUMD, J_TOTALDEBIT=J_TOTALDEBIT-:DD, '
                +'J_TOTALCREDIT=J_TOTALCREDIT-:CD, J_TOTDEBE=J_TOTDEBE-:DE, J_TOTCREE=J_TOTCREE-:CE,'
                +'J_TOTDEBS=J_TOTDEBS-:DS, J_TOTCRES=J_TOTCRES-:CS, '
                +'J_TOTDEBP=J_TOTDEBP-:DP, J_TOTCREP=J_TOTCREP-:CP WHERE J_JOURNAL=:CPTE' ;
   end ;
Q.SQL.Add(SQL) ; ChangeSQL(Q) ; PrepareSQLODBC(Q) ;
Result:=Q ;
end ;
{$ENDIF}

//=======================================================
//================== Génération des RB ==================
//=======================================================
function TZBalance.RbDel : Boolean ;
var QEcr, QAna : TQuery ; Where, JalRb, Cols, WhereAna, ColsAna : string ;
    JalSoldeD, JalSoldeC : Double ; bEcr, bAna : Boolean ;
    lFRM : TFRM;
    ll : LongInt;
begin
  Result := TRUE ;
  bAna   := False ;
  JalRb  := '' ;

  Where := ' E_QUALIFORIGINE="'+FOpt.TypeBal+'"'
          +' AND E_EXERCICE="'+FOpt.ExoBal+'"'
          +' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour))+'"'
          +' AND E_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"'
          +' AND E_CREERPAR<>"DET"' ;
  Cols:='E_GENERAL, E_AUXILIAIRE, E_DEBIT, E_CREDIT, E_JOURNAL, E_NUMEROPIECE' ;

  QEcr:=OpenSQL('SELECT '+Cols+' FROM ECRITURE WHERE'+Where, TRUE) ;
  bEcr:=not QEcr.EOF ;
  if bEcr then
  begin
    JalRb:=QEcr.FindField('E_JOURNAL').AsString ;
    FNumPiece:=QEcr.FindField('E_NUMEROPIECE').AsInteger ;
  end
  else
    FNumPiece:=-1 ;

  // Comptes généraux
  JalSoldeD := 0 ;
  JalSoldeC := 0 ;
  while not QEcr.EOF do
  begin
    Fillchar(lFRM, SizeOf(lFRM),#0);
    lFRM.Cpt := QEcr.FindField('E_GENERAL').AsString ;
    AttribParamsNew(lFRM,
                    QEcr.FindField('E_DEBIT').AsFloat,
                    QEcr.FindField('E_CREDIT').AsFloat,
                    FOpt.TypeExo);

    JalSoldeD := JalSoldeD + QEcr.FindField('E_DEBIT').AsFloat ;
    JalSoldeC := JalSoldeC + QEcr.FindField('E_CREDIT').AsFloat ;

    ll := ExecReqINVNew( fbGene, lFRM);
    if ll <> 1 then
    begin
      Result := False;
      V_PGI.IoError := oeSaisie ;
    end;
    QEcr.Next;
  end;

  // Comptes tiers
  if bEcr then QEcr.First ;  // CA - 03/10/2002 - erreur SQL si QEcr vide
  while not QEcr.EOF do
  begin
    if QEcr.FindField('E_AUXILIAIRE').AsString = '' then
    begin
      QEcr.Next ;
      Continue ;
    end ;
    Fillchar(lFRM, SizeOf(lFRM),#0);
    lFRM.Cpt := QEcr.FindField('E_AUXILIAIRE').AsString ;
    AttribParamsNew(lFRM,
                    QEcr.FindField('E_DEBIT').AsFloat,
                    QEcr.FindField('E_CREDIT').AsFloat,
                    FOpt.TypeExo);

    ll := ExecReqINVNew( fbAux, lFRM);
    if ll <> 1 then
    begin
      Result := False;
      V_PGI.IoError := oeSaisie ;
    end;
    QEcr.Next ;
  end ;

  // Sections
  if bEcr then
  begin
   {$IFDEF EAGLCLIENT}
      WhereAna := ' Y_NUMEROPIECE='+ IntToStr(QEcr.Detail[QEcr.Detail.Count-1].GetValue('E_NUMEROPIECE'))
   {$ELSE}
      WhereAna := ' Y_NUMEROPIECE='+ IntToStr(QEcr.FindField('E_NUMEROPIECE').AsInteger)
   {$ENDIF}
               +' AND Y_JOURNAL="'+JalRb+'"'
               +' AND Y_EXERCICE="'+FOpt.ExoBal+'"'
               +' AND Y_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour))+'"'
               +' AND Y_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"' ;

    ColsAna := 'Y_SECTION, Y_AXE, Y_DEBIT, Y_CREDIT' ;
    QAna := nil;
    try
      try
        QAna := OpenSQL('SELECT ' + ColsAna + ' FROM ANALYTIQ WHERE ' + WhereAna, TRUE) ;
        bAna := not QAna.EOF ;
        while not QAna.EOF do
        begin
          Fillchar(lFRM, SizeOf(lFRM),#0);
          lFRM.Cpt := QAna.FindField('Y_SECTION').AsString ;
          lFRM.Axe := QAna.FindField('Y_AXE').AsString ;
          AttribParamsNew(lFRM,
                          QEcr.FindField('Y_DEBIT').AsFloat,
                          QEcr.FindField('Y_CREDIT').AsFloat,
                          FOpt.TypeExo);

          ll := ExecReqINVNew( fbSect, lFRM);
          if ll <> 1 then
          begin
            Result := False;
            V_PGI.IoError := oeSaisie ;
          end;
          QAna.Next;
        end ;
      except
        on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : RbDel -> Section');
      end;
    finally
      Ferme(QAna) ;
    end;
  end
  else
    bAna := False;

  Ferme(QEcr) ;

  // Le journal
  if bEcr then
  begin
    Fillchar(lFRM, SizeOf(lFRM),#0);
    lFRM.Cpt := JalRb ;
    AttribParamsNew(lFRM, JalSoldeD, JalSoldeC, FOpt.TypeExo) ;
    ll := ExecReqINVNew( fbJal, lFRM);
    if ll <> 1 then
    begin
      Result := False;
      V_PGI.IoError := oeSaisie ;
    end;
  end ;

  // Destruction des écritures
  if bEcr then
     if ExecuteSQL('DELETE FROM ECRITURE WHERE'+Where)<=0 then
     begin
       Result := FALSE ;
       V_PGI.IOError := oeSaisie ;
       Exit ;
     end ;

  // Destruction des écritures analytiques
  if bAna then
     if ExecuteSQL('DELETE FROM ANALYTIQ WHERE'+WhereAna)<=0 then
     begin
       Result := FALSE ;
       V_PGI.IOError := oeSaisie ;
       Exit ;
     end ;
end ;

function TZBalance.RbRead : Boolean ;
var QEcr : TQuery ; Cols, Where : string ;
begin
Cols:='*' ;
Where:=' E_EXERCICE="'+FOpt.ExoBal+'"'
      +' AND E_QUALIFORIGINE="'+FOpt.TypeBal+'"'
      +' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour))+'"'
      +' AND E_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"'
      +' AND E_NUMEROPIECE='+IntToStr(FOpt.NumPiece) ;
QEcr:=OpenSQL('SELECT '+Cols+' FROM ECRITURE WHERE'+Where, TRUE) ;
Result:=not QEcr.EOF ;
while not QEcr.EOF do
  begin
  // Charger le TZF
  CreateRow(QEcr, Ecrs.Detail.Count+1) ;
  QEcr.Next ;
  end ;
Ferme(QEcr) ;
end ;

function TZBalance.RbPrepare(EcrMaster : TZF; JalRb, sMode: string; DateEcr: TDateTime;
                             NumPiece: Integer; var k: Integer; SOpp : Double; bEcart :Boolean) : TZF ;
var Q : TQuery ; EcrRb : TZF ; z : Integer ; Solde : Double ;
    Mode, DecEnc : string ; bLe, bPo : Boolean ;
begin
EcrRb:=TZF.Create('ECRITURE', nil, -1) ;
EcrRb.SetDateModif(Date) ;
for z:=1 to MAXAXE do TZF.Create('A'+IntToStr(z), EcrRb, -1) ;
EcrRb.PutValue('E_EXERCICE',      FOpt.ExoBal) ;
EcrRb.PutValue('E_JOURNAL',       JalRb) ;
EcrRb.PutValue('E_DATECOMPTABLE', DateEcr) ;
EcrRb.PutValue('E_NUMEROPIECE',   NumPiece) ;
EcrRb.PutValue('E_NUMLIGNE',      k); k:=k+1 ;
EcrRb.PutValue('E_GENERAL',       EcrMaster.GetValue('HB_COMPTE1')) ;
EcrRb.PutValue('E_AUXILIAIRE',    EcrMaster.GetValue('HB_COMPTE2')) ;
EcrRb.PutValue('E_LIBELLE',       EcrMaster.GetValue('LIBELLE')) ;
EcrRb.PutValue('E_DEBIT',         SoldeD('HB_DEBIT', 'HB_CREDIT', EcrMaster, nil)) ;
EcrRb.PutValue('E_CREDIT',        SoldeC('HB_DEBIT', 'HB_CREDIT', EcrMaster, nil)) ;
EcrRb.PutValue('E_DEBITDEV',      SoldeD('HB_DEBITDEV', 'HB_CREDITDEV', EcrMaster, nil)) ;
EcrRb.PutValue('E_CREDITDEV',     SoldeC('HB_DEBITDEV', 'HB_CREDITDEV', EcrMaster, nil)) ;
EcrRb.PutValue('E_NATUREPIECE',   'OD') ;
EcrRb.PutValue('E_TYPEMVT',       'DIV') ;
EcrRb.PutValue('E_QUALIFPIECE',   'N') ;
EcrRb.PutValue('E_VALIDE',        '-') ;
EcrRb.PutValue('E_DEVISE',        EcrMaster.GetValue('HB_DEVISE')) ;
EcrRb.PutValue('E_TAUXDEV',       1) ;
EcrRb.PutValue('E_QUALIFORIGINE', FOpt.TypeBal) ;
EcrRb.PutValue('E_ECRANOUVEAU',   'N') ;
EcrRb.PutValue('E_CREERPAR',      'BAL') ;
EcrRb.PutValue('E_MODESAISIE',    sMode) ;
EcrRb.PutValue('E_PERIODE',       GetPeriode(DateEcr)) ;
EcrRb.PutValue('E_SEMAINE',       NumSemaine(DateEcr)) ;
bLe:=(Copy(EcrMaster.GetValue('LETTRABLE'), 1, 1)='X') ;
bPo:=(Copy(EcrMaster.GetValue('LETTRABLE'), 2, 1)='X') ;
if (bLe) or (bPo) then
  begin
  Mode:='' ;
  Q:=OpenSQL('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="'
             +EcrMaster.GetValue('MODEREGL')+'"', TRUE) ;
  if not Q.EOF then Mode:=Q.FindField('MR_MP1').AsString ;
  Ferme(Q) ;
  if Mode='' then
    begin
    Q:=OpenSQL('SELECT MP_MODEPAIE FROM MODEPAIE', TRUE) ;
    if not Q.EOF then Mode:=Q.FindField('MP_MODEPAIE').AsString ;
    Ferme(Q) ;
    end ;
  EcrRb.PutValue('E_MODEPAIE',        Mode) ;
  EcrRb.PutValue('E_DATEECHEANCE',    DateEcr) ;
  EcrRb.PutValue('E_ORIGINEPAIEMENT', DateEcr) ;
  EcrRb.PutValue('E_ECHE',            'X') ;
  EcrRb.PutValue('E_NUMECHE',         1) ;
  if bLe then
    begin
    EcrRb.PutValue('E_ETATLETTRAGE',       'AL') ;
    Solde:=EcrRb.GetValue('E_DEBIT')-EcrRb.GetValue('E_CREDIT') ;
    if Solde=0 then DecEnc:='' ;
    if Solde>0 then DecEnc:='ENC' else DecEnc:='DEC' ;
    EcrRb.PutValue('E_ENCAISSEMENT', DecEnc) ;
    EcrRb.PutValue('E_RIB', GetRIBPrincipal(EcrMaster.GetValue('HB_COMPTE2'))) ;
    end ;
  EcrRb.PutValue('E_DATEVALEUR', DateEcr) ;
  end ;
if VentilerTOB(EcrRb, '', 0, FDev.Decimale, FALSE) then EcrRb.PutValue('E_ANA', 'X') ;
EcrRb.InsertDB(nil) ;
Result:=EcrRb ;
end ;

function TZBalance.RbWrite : Boolean ;
var Q :TQuery ; EcrRb, EcrMaster, ee : TZF ; e : TList ; JalRb, Critere : string ;
    JalSoldeD, JalSoldeC : Double ; i, j, k, c, NumPiece : Integer ;
    DateEcr : TDateTime ; sMode, sCompteur : string ;
    lFRM : TFRM;
    ll : LongInt;
    //SEuro, TDEuro, TCEuro : Double ;
begin
Result:=TRUE ; sMode:='' ;
DateEcr:=EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour) ;
// Journal des OD
JalRb:=VH^.JalRepBalAN ;
if JalRb='' then
  begin
  Q:=OpenSQL('SELECT J_JOURNAL, J_MODESAISIE, J_COMPTEURNORMAL FROM JOURNAL WHERE J_NATUREJAL="OD" ORDER BY J_JOURNAL', TRUE) ;
  JalRb:=Q.Fields[0].AsString ;
  sMode:=Q.Fields[1].AsString ;
  sCompteur:=Q.Fields[2].AsString ;
  Ferme(Q) ;
  end else
  begin
  Q:=OpenSQL('SELECT J_MODESAISIE, J_COMPTEURNORMAL FROM JOURNAL WHERE J_JOURNAL="'+JalRb+'" ORDER BY J_JOURNAL', TRUE) ;
  sMode:=Q.Fields[0].AsString ;
  sCompteur:=Q.Fields[1].AsString ;
  Ferme(Q) ;
  end ;
// Numéro de pièce à utiliser
if FNumPiece<0 then
  begin
  if (sMode<>'') and (sMode<>'-') then
    begin
    Critere:='E_JOURNAL="'+JalRb+'" AND E_NUMLIGNE=1' ;
    Q:=OpenSQL('SELECT DISTINCT E_NUMEROPIECE FROM ECRITURE WHERE '+Critere
              +' ORDER BY E_NUMEROPIECE DESC', TRUE) ;
    FNumPiece:=Q.Fields[0].AsInteger+1 ;
    NumPiece := FNumPiece;  // CA - 06/03/2002 - c'est NumPiece qui est utilisé.
    Ferme(Q) ;
    end else
    begin
    SetIncNum(EcrGen, sCompteur, NumPiece, DateEcr) ;
    end ;
  end else NumPiece:=FNumPiece ;
// Let's go !!!
InitMove(Ecrs.Detail.Count, TraduireMemoire('Génération des écritures')) ;
k:=1 ; e:=TList.Create ;
for i:=0 to Ecrs.Detail.Count-1 do
    begin
    MoveCur(FALSE) ;
    EcrMaster:=TZF(Ecrs.Detail[i]) ;
    if EcrMaster.GetValue('BADROW')='X' then Continue ;
    if EcrMaster.Detail.Count>0 then
      begin
      for j:=0 to EcrMaster.Detail.Count-1 do
        begin
        EcrRb:=RbPrepare(TZF(EcrMaster.Detail[j]), JalRb, sMode, DateEcr, NumPiece, k) ;
        e.Add(EcrRb) ;
        end ;
      end else
      begin
      EcrRb:=RbPrepare(EcrMaster, JalRb, sMode, DateEcr, NumPiece, k) ;
      e.Add(EcrRb) ;
      end ;
    end ;
FiniMove ;

  // Les comptes géréraux
  JalSoldeD := 0 ;
  JalSoldeC := 0 ;
  InitMove(e.Count, TraduireMemoire('Mise à jour des comptes généraux')) ;
  for i:=0 to e.Count-1 do
  begin
    MoveCur(FALSE) ;
    Fillchar(lFRM, SizeOf(lFRM),#0);
    lFRM.cpt   := TZF(e[i]).GetValue('E_GENERAL') ;
    lFRM.NumD  := TZF(e[i]).GetValue('E_NUMEROPIECE') ;
    lFRM.DateD := TZF(e[i]).GetValue('E_DATECOMPTABLE') ;
    lFRM.LigD  := TZF(e[i]).GetValue('E_NUMLIGNE') ;

    JalSoldeD := JalSoldeD+TZF(e[i]).GetValue('E_DEBIT') ;
    JalSoldeC := JalSoldeC+TZF(e[i]).GetValue('E_CREDIT') ;

    AttribParamsNew(lFRM, TZF(e[i]).GetValue('E_DEBIT'),
                    TZF(e[i]).GetValue('E_CREDIT'), FOpt.TypeExo) ;

    ll := ExecReqMAJ( fbGene, False, False, lFRM ) ;
    if ll <> 1 then
    begin
      Result := False;
      V_PGI.IoError := oeSaisie ;
    end;
  end ;
  FiniMove ;
  
  // Les comptes tiers
  InitMove(e.Count, TraduireMemoire('Mise à jour des comptes tiers')) ;
  for i:=0 to e.Count-1 do
  begin
    MoveCur(FALSE) ;
    if TZF(e[i]).GetValue('E_AUXILIAIRE') = '' then Continue ;
    Fillchar(lFRM, SizeOf(lFRM),#0);
    lFRM.Cpt   := TZF(e[i]).GetValue('E_AUXILIAIRE') ;
    lFRM.NumD  := TZF(e[i]).GetValue('E_NUMEROPIECE') ;
    lFRM.DateD := TZF(e[i]).GetValue('E_DATECOMPTABLE') ;
    lFRM.LigD  := TZF(e[i]).GetValue('E_NUMLIGNE') ;

    AttribParamsNew(lFRM, TZF(e[i]).GetValue('E_DEBIT'),
                    TZF(e[i]).GetValue('E_CREDIT'), FOpt.TypeExo) ;

    ll := ExecReqMAJ( fbAux, False, False, lFRM ) ;
    if ll <> 1 then
    begin
      Result := False;
      V_PGI.IoError := oeSaisie ;
    end;
  end;
  FiniMove ;
  
  // Sections
  InitMove(e.Count, TraduireMemoire('Mise à jour des sections')) ;
  //Q:=FabricReqMAJ(fbSect, FALSE, FALSE) ;
  for i:=0 to e.Count-1 do
  begin
    MoveCur(FALSE) ;
    if TZF(e[i]).GetValue('E_ANA')<>'X' then Continue ;
    for j:=0 to MAXAXE-1 do
    begin
      c := TZF(e[i]).Detail[j].Detail.Count ;
      if c = 0 then Continue ;
      for k := 0 to c-1 do
      begin
        ee:=TZF(TZF((TZF(e[i]).Detail[j]).Detail[k])) ;
        Fillchar(lFRM, SizeOf(lFRM),#0);
        lFRM.Cpt   := ee.GetValue('Y_SECTION') ;
        lFRM.Axe   := ee.GetValue('Y_AXE') ;
        lFRM.Numd  := TZF(e[i]).GetValue('E_NUMEROPIECE') ;
        lFRM.DateD := TZF(e[i]).GetValue('E_DATECOMPTABLE') ;
        lFRM.LigD  := TZF(e[i]).GetValue('E_NUMLIGNE') ;

        AttribParamsNew(lFRM, ee.GetValue('Y_DEBIT'), ee.GetValue('Y_CREDIT'), FOpt.TypeExo) ;

        ll := ExecReqMAJ( fbSect, False, False, lFRM ) ;
        if ll <> 1 then
        begin
          Result := False;
          V_PGI.IoError := oeSaisie ;
        end;
      end ;
    end ;
  end ;
  FiniMove ;
  
  // Le journal
  Fillchar(lFRM, SizeOf(lFRM),#0);
  lFRM.Cpt   := JalRb ;
  lFRM.Numd  := NumPiece ;
  lFRM.DateD := EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour) ;
  AttribParamsNew(lFRM, JalSoldeD, JalSoldeC, FOpt.TypeExo) ;
  ll := ExecReqMAJ( fbJal, False, False, lFRM ) ;
  if ll <> 1 then
  begin
    Result := False;
    V_PGI.IoError := oeSaisie ;
  end;

  VideListe(e) ;
  e.Free ;
  FNumPiece := -1 ;
// SO_EXOV8
//if ExecuteSQL('UPDATE SOCIETE SET SO_EXOV8="'+FOpt.ExoAno+'"')<=0 then
//   begin Result:=FALSE ; V_PGI.IOError:=oeSaisie ; Exit ; end ;
end ;

//=======================================================
//=============== Génération Automatique ================
//=======================================================
// SELECT E_GENERAL, SUM(E_DEBIT), SUM(E_CREDIT) FROM ECRITURE GROUP BY E_GENERAL ORDER BY E_GENERAL
// SELECT E_GENERAL, E_AUXILIAIRE, SUM(E_DEBIT), SUM(E_CREDIT) FROM ECRITURE GROUP BY E_GENERAL, E_AUXILIAIRE ORDERBY E_GENERAL, E_AUXILIAIRE
// SELECT G_GENERAL, G_LIBELLE, G_TOTALDEBIT, G_TOTALCREDIT FROM GENERAUX WHERE G_TOTALDEBIT<>0 OR G_TOTALCREDIT<>0 ORDER BY E_GENERAL
function TZBalance.GenAuto : Boolean ;
var Q, QGen : TQuery ; Ecr : TZF ; Plan : TZPlan ; k :Integer ;
    Cols, Where, GroupBy, OrderBy, NumCompte : string ; nRow : LongInt ;
begin
nRow:=0 ;
// Charger les collectifs
if not FOpt.WithAux then
  begin Plan:=TZPlan.Create ; Plan.Load(FALSE, FALSE, TRUE) ; end else Plan:=nil ;
if FOpt.WithAux then Cols:='E_GENERAL, SUM(E_DEBIT), SUM(E_CREDIT), E_AUXILIAIRE'
                else Cols:='E_GENERAL, SUM(E_DEBIT), SUM(E_CREDIT)' ;
Where:='E_QUALIFPIECE="N"'
      +' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour))+'"'
      +' AND E_DATECOMPTABLE<="'+USDateTime(FinDeMois(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour)))+'"' ;
if FOpt.WithAux then GroupBy:='E_GENERAL, E_AUXILIAIRE' else GroupBy:='E_GENERAL' ;
OrderBy:=GroupBy ;
Q:=OpenSQL('SELECT '+Cols+' FROM ECRITURE WHERE '+Where+' GROUP BY '+GroupBy+' ORDER BY '+OrderBy, TRUE) ;
Result:=not Q.EOF ;
while not Q.EOF do
  begin
  // Si Débit et Crédit à 0, je ne génère pas de ligne !!!
  if (Q.Fields[1].AsFloat=0) and (Q.Fields[2].AsFloat=0) then begin Q.Next ; Continue ; end ;
  QGen:=OpenSQL('SELECT * FROM HISTOBAL WHERE HB_TYPEBAL="'+W_W+'"', FALSE) ;
  QGen.Insert ; InitNew(QGen) ;
  Ecr:=CreateRow(nil, nRow) ;
  QGen.FindField('HB_TYPEBAL').AsString:=FOpt.TypeBal ;
    Ecr.PutValue('HB_TYPEBAL', FOpt.TypeBal) ;
  QGen.FindField('HB_PLAN').AsString:=FOpt.PlanBal ;
    Ecr.PutValue('HB_PLAN', FOpt.PlanBal) ;
  QGen.FindField('HB_EXERCICE').AsString:=FOpt.ExoBal ;
    Ecr.PutValue('HB_EXERCICE', FOpt.ExoBal) ;
  QGen.FindField('HB_DATE1').AsDateTime:=EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour) ;
    Ecr.PutValue('HB_DATE1', EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour)) ;
  QGen.FindField('HB_DATE2').AsDateTime:=EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour) ;
    Ecr.PutValue('HB_DATE2', EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour)) ;
  QGen.FindField('HB_COMPTE1').AsString:=Q.FindField('E_GENERAL').AsString ;
    Ecr.PutValue('HB_COMPTE1', Q.FindField('E_GENERAL').AsString) ;
  if FOpt.WithAux then
    begin
    QGen.FindField('HB_COMPTE2').AsString:=Q.FindField('E_AUXILIAIRE').AsString ;
      Ecr.PutValue('HB_COMPTE2', Q.FindField('E_AUXILIAIRE').AsString) ;
    end else
    begin
    NumCompte:=Q.FindField('E_GENERAL').AsString ;
    k:=Plan.FindCompte(NumCompte) ;
    if k>=0 then
      begin
      if Plan.GetValue('G_NATUREGENE', k)='COC' then QGen.FindField('HB_COMPTE2').AsString:=VH^.TiersDefCli else
        if Plan.GetValue('G_NATUREGENE', k)='COF' then QGen.FindField('HB_COMPTE2').AsString:=VH^.TiersDefFou else
          if Plan.GetValue('G_NATUREGENE', k)='COS' then QGen.FindField('HB_COMPTE2').AsString:=VH^.TiersDefSal else
            QGen.FindField('HB_COMPTE2').AsString:=VH^.TiersDefDiv ;
      Ecr.PutValue('HB_COMPTE2', QGen.FindField('HB_COMPTE2').AsString) ;
      end ;
    end ;
  if (Q.Fields[1].AsFloat>0) then
    begin
    QGen.FindField('HB_DEBIT').AsFloat:=Q.Fields[1].AsFloat ;
      Ecr.PutValue('HB_DEBIT', Q.Fields[1].AsFloat) ;
    end ;
  if (Q.Fields[2].AsFloat>0) then
    begin
    QGen.FindField('HB_CREDIT').AsFloat:=Abs(Q.Fields[2].AsFloat) ;
      Ecr.PutValue('HB_CREDIT', Abs(Q.Fields[2].AsFloat)) ;
    end ;
  Ecr.SetMontants(Ecr.GetValue('HB_DEBIT'), Ecr.GetValue('HB_CREDIT'),
                  FDev, FALSE) ;
  QGen.FindField('HB_DEBITDEV').AsFloat:=Ecr.GetValue('HB_DEBITDEV') ;
  QGen.FindField('HB_CREDITDEV').AsFloat:=Ecr.GetValue('HB_CREDITDEV') ;
  QGen.FindField('HB_DEVISE').AsString:=FDev.Code ;
  QGen.FindField('HB_ETABLISSEMENT').AsString:=Ecr.GetValue('HB_ETABLISSEMENT') ;
  QGen.FindField('HB_SOCIETE').AsString:=Ecr.GetValue('HB_SOCIETE') ;
  QGen.Post ;
  nRow:=nRow+1 ;
  Ferme(QGen) ;
  Q.Next ;
  end ;
Ferme(Q) ;
Plan.Free ;
end ;

function TZBalance.GenAuto2 : Boolean ;
var Q : TQuery ; Ecr : TZF ; Plan : TZPlan ; k :Integer ;
    Cols, Where, GroupBy, OrderBy, NumCompte : string ; nRow : LongInt ;
    LastCompte, LastAux : string ;
    bAno : boolean;
begin
  nRow:=0 ; Ecr := nil;
  LastCompte:='' ; LastAux:='' ;
  // Charger les collectifs
  if not FOpt.WithAux then
  begin
    Plan:=TZPlan.Create ;
    Plan.Load(FALSE, FALSE, TRUE) ;
  end else Plan:=nil ;
  if FOpt.WithAux then Cols:='E_GENERAL, E_ECRANOUVEAU, SUM(E_DEBIT)-SUM(E_CREDIT), SUM(E_DEBIT), SUM(E_CREDIT), E_AUXILIAIRE'
    else Cols:='E_GENERAL, E_ECRANOUVEAU, SUM(E_DEBIT)-SUM(E_CREDIT), SUM(E_DEBIT), SUM(E_CREDIT)' ;
  Where:='E_QUALIFPIECE="N"'
      +' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour))+'"'
      +' AND E_DATECOMPTABLE<="'+USDateTime(FinDeMois(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour)))+'"' ;
  if FOpt.WithAux then GroupBy:='E_GENERAL, E_AUXILIAIRE, E_ECRANOUVEAU'
    else GroupBy:='E_GENERAL, E_ECRANOUVEAU' ;
  OrderBy:=GroupBy ;
  Q:=OpenSQL('SELECT '+Cols+' FROM ECRITURE WHERE '+Where+' GROUP BY '+GroupBy+' ORDER BY '+OrderBy, TRUE) ;
  Result:=not Q.EOF ;
  while not Q.EOF do
  begin
    bAno:=(Q.Fields[1].AsString<>'N') ;
    // Si Débit et Crédit à 0, je ne génère pas de ligne !!!
    if (Q.Fields[2].AsFloat=0) and (Q.Fields[3].AsFloat=0) then
    begin
      Q.Next ;
      Continue ;
    end ;
    if ((not bAno) or
      ((bAno) and (LastCompte<>Q.FindField('E_GENERAL').AsString)) or
      ((bAno) and (FOpt.WithAux) and
      ((LastCompte<>Q.FindField('E_GENERAL').AsString)) and
      ((LastAux<>Q.FindField('E_AUXILIAIRE').AsString)))) then
    begin
      Ecr:=CreateRow(nil, nRow) ;
      LastCompte:=Q.FindField('E_GENERAL').AsString ;
      if FOpt.WithAux then LastAux:=Q.FindField('E_AUXILIAIRE').AsString ;
    end ;
    if Ecr <> nil then
    begin
      Ecr.PutValue('HB_TYPEBAL', FOpt.TypeBal) ;
      Ecr.PutValue('HB_PLAN', FOpt.PlanBal) ;
      Ecr.PutValue('HB_EXERCICE', FOpt.ExoBal) ;
      Ecr.PutValue('HB_DATE1', EncodeDate(FOpt.DebYear, FOpt.DebMonth, FOpt.DebJour)) ;
      Ecr.PutValue('HB_DATE2', EncodeDate(FOpt.Year,    FOpt.Month,    FOpt.MaxJour)) ;
      Ecr.PutValue('HB_COMPTE1', Q.FindField('E_GENERAL').AsString) ;
      if FOpt.WithAux then Ecr.PutValue('HB_COMPTE2', Q.FindField('E_AUXILIAIRE').AsString)
      else
      begin
        NumCompte:=Q.FindField('E_GENERAL').AsString ;
        k:=Plan.FindCompte(NumCompte) ;
        if k>=0 then
        begin
          if Plan.GetValue('G_NATUREGENE', k)='COC' then Ecr.PutValue('HB_COMPTE2', VH^.TiersDefCli) else
            if Plan.GetValue('G_NATUREGENE', k)='COF' then Ecr.PutValue('HB_COMPTE2', VH^.TiersDefFou) else
              if Plan.GetValue('G_NATUREGENE', k)='COS' then Ecr.PutValue('HB_COMPTE2', VH^.TiersDefSal) else
                Ecr.PutValue('HB_COMPTE2', VH^.TiersDefDiv) ;
        end ;
      end ;
      if bAno then
      begin
        if (Arrondi(Q.Fields[2].AsFloat, FDev.Decimale)>0) then
          Ecr.PutValue('HB_DEBIT',  Ecr.GetValue('HB_DEBIT')+Arrondi(Q.Fields[2].AsFloat, FDev.Decimale))
        else Ecr.PutValue('HB_CREDIT', Ecr.GetValue('HB_CREDIT')+Abs(Arrondi(Q.Fields[2].AsFloat, FDev.Decimale))) ;
      end else
      begin
        if (Q.Fields[3].AsFloat>0) then Ecr.PutValue('HB_DEBIT',  Q.Fields[3].AsFloat) ;
        if (Q.Fields[4].AsFloat>0) then Ecr.PutValue('HB_CREDIT', Q.Fields[4].AsFloat) ;
      end ;
      Ecr.SetMontants(Ecr.GetValue('HB_DEBIT'), Ecr.GetValue('HB_CREDIT'),
                    FDev, FALSE) ;
    end;
    nRow:=nRow+1 ;
    Q.Next ;
  end ;
  Ferme(Q) ;
  Plan.Free ;
  Ecrs.InsertDB(nil, TRUE) ;
end ;

end.
