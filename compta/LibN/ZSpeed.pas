{***********UNITE*************************************************
Auteur  ...... : Pascal FUGIER
Créé le ...... : 21/02/2000
Modifié le ... : 25/02/2000
Description .. : Classe ZSpeed qui permet d'optimiser les fonctions @
Suite ........ : déclenchées par le générateur d'état
Suite ........ : 
Suite ........ : C'est d'la BOMBE
Mots clefs ... : 
*****************************************************************}
unit ZSpeed;

// MODIF ENT1
//Type LaVariableHalley = RECORD
//(*======= PFUGIER *)
//     SpeedObj : Pointer ;
//     bByGenEtat : Boolean ;
//(*======= PFUGIER *)
//     END ;

//Procedure InitLaVariableHalley ;
//BEGIN
//New(VH) ; FillChar(VH^,Sizeof(VH^),#0) ;
//VH^.ImportRL:=FALSE ; VH^.TenueEuro:=False ; VH^.Mugler:=FALSE ; VH^.MultiSouche:=FALSE ;
//VH^.PasParSoc:=False ; VH^.BouclerSaisieCreat:=True ; VH^.AttribRIBAuto:=False ;
//VH^.SpeedCum:=FALSE ;
//(*======= PFUGIER *)
//VH^.bByGenEtat:=FALSE ; VH^.SpeedObj:=nil ;
//(*======= PFUGIER *)

interface

uses Forms, Classes, UTob, Ent1, CpteUtil, tCalcCum ;

type TZSpeed = class
     private
     ECacheTOB:  TList ;
     ECachePAR:  TList ;
     HCacheTOB:  TList ;
     HCachePAR:  TList ;
     bSQLRub:    Boolean ;
     Rub:        TOB ;
     bAnalyse:   Boolean ;
     Critere:    string ;
     procedure   EInitCache ;
     procedure   HInitCache ;
     procedure   ECleanUp ;
     procedure   HCleanUp ;
     procedure   RCleanUp ;
     function    ESetParams(zParams: ZSParams): Integer ;
     function    EVerifParams(zParams: ZSParams): Integer ;
     function    HSetParams(zParams: ZSParams): Integer ;
     function    HVerifParams(zParams: ZSParams): Integer ;
     procedure   EGetSQLCumul(iCache: Integer) ;
     procedure   HGetSQLCumul(iCache: Integer) ;
     procedure   GetSQLRub(sCodeRub: string) ;
     procedure   CalculDC(iCache: Integer; P: string; T: TOB; var lCalc: TabTot; nbDec: Integer; bHisto: Boolean) ;
     procedure   Calcul(var zCalc: TabTot; lCalc: TabTot; NbDec: Integer; stSolde, Signe: string) ;
     procedure   Parser(iCache: Integer; itParse, P: string; SaufL: TStringList; var zCalc: TabTot; nbDec: Integer; Signe, Formule: string; bHisto: Boolean) ;
     public
     constructor Create(AOwner: TComponent) ;
//     constructor CreateP(AOwner: TComponent; zParams: ZSParams) ;
     destructor  Destroy ; override ;
     function    GetResult(Total : TabTot; SetTyp : SetttTypePiece; Ano, Signe: string) : Double ;
     function    GetCumul(Cpt, Sauf : string; zParams: ZSParams; NbDec: Integer; Signe, Formule: string): TabTot ;
     function    GetRub(CodeRub: string): TOB ;
     procedure   InitAnalyseCompte(CritDeb, CritFin: string) ;
     end ;

implementation

uses HEnt1,
{$IFDEF EAGLCLIENT}

{$ELSE}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls, ParamSoc ;

const CACHE_MAX = 3 ;

constructor TZSpeed.Create(AOwner: TComponent) ;
begin
ECacheTOB:=nil ;
ECachePAR:=nil ;
HCacheTOB:=nil ;
HCachePAR:=nil ;
Rub:=nil ;
bSQLRub:=FALSE ;
bAnalyse:=FALSE ;
Critere:='' ;
end ;

(*
constructor TZSpeed.CreateP(AOwner: TComponent; zParams: ZSParams) ;
begin
ECacheTOB:=nil ;
ECachePAR:=nil ;
HCacheTOB:=nil ;
HCachePAR:=nil ;
bSQLRub:=FALSE ;   Rub:=nil ;
EInitCache ;
HInitCache ;
SetParams(zParams) ;
end ;
*)

destructor TZSpeed.Destroy ;
begin
ECleanUp ;
HCleanUp ;
RCleanUp ;
end ;

procedure TZSpeed.InitAnalyseCompte(CritDeb, CritFin: string) ;
begin
if (CritDeb='') and (CritFin='') then begin bAnalyse:=FALSE ; Critere:='' ; Exit ; end ;
if (CritDeb='') and (CritFin<>'') then Critere:=CritFin+';' else
  if (CritDeb<>'') and (CritFin='') then Critere:=CritDeb+';' else
    if (CritDeb<>'') and (CritFin<>'') then Critere:=CritDeb+':'+CritFin+';' ;
bAnalyse:=TRUE ;
end ;

procedure TZSpeed.EInitCache ;
begin
ECacheTOB:=TList.Create ;
ECachePAR:=TList.Create ;
end ;

procedure TZSpeed.HInitCache ;
begin
HCacheTOB:=TList.Create ;
HCachePAR:=TList.Create ;
end ;

procedure TZSpeed.ECleanUp ;
var Params: PZSParams ; T: TOB ;
begin
if ECacheTOB<>nil then
  begin
  while ECacheTOB.Count>0 do
    begin T:=ECacheTOB.Items[0] ; T.Free ; ECacheTOB.Delete(0) ; end ;
  ECacheTOB.Free ; ECacheTOB:=nil ;
  end ;
if ECachePAR<>nil then
  begin
  while ECachePAR.Count>0 do
    begin Params:=ECachePAR.Items[0] ; Dispose(Params) ; ECachePAR.Delete(0) ; end ;
  ECachePAR.Free ; ECachePAR:=nil ;
  end ;
end ;

procedure TZSpeed.HCleanUp ;
var Params: PZSParams ; T: TOB ;
begin
if HCacheTOB<>nil then
  begin
  while HCacheTOB.Count>0 do
    begin T:=HCacheTOB.Items[0] ; T.Free ; HCacheTOB.Delete(0) ; end ;
  HCacheTOB.Free ; HCacheTOB:=nil ;
  end ;
if HCachePAR<>nil then
  begin
  while HCachePAR.Count>0 do
    begin Params:=HCachePAR.Items[0] ; Dispose(Params) ; HCachePAR.Delete(0) ; end ;
  HCachePAR.Free ; HCachePAR:=nil ;
  end ;
end ;

procedure TZSpeed.RCleanUp ;
begin
if Rub<>nil then begin Rub.Free ; Rub:=nil ; end ;
end ;

function TZSpeed.ESetParams(zParams: ZSParams): Integer ;
var Params: PZSParams ;
begin
if (ECachePAR=nil) and (ECacheTOB=nil) then EInitCache ;
New(Params) ;
Params^:=zParams ;
Params^.bSQLCumul:=FALSE ;
Result:=ECachePAR.Add(Params) ;
end ;

function TZSpeed.HSetParams(zParams: ZSParams): Integer ;
var Params: PZSParams ;
begin
if (HCachePAR=nil) and (HCacheTOB=nil) then HInitCache ;
New(Params) ;
Params^:=zParams ;
Params^.bSQLCumul:=FALSE ;
Result:=HCachePAR.Add(Params) ;
end ;

function TZSpeed.EVerifParams(zParams: ZSParams): Integer ;
var Params: PZSPARAMS ; i: Integer ; bMatch: Boolean ;
begin
Result:=-1 ;
for i:=0 to ECachePAR.Count-1 do
  begin
  bMatch:=TRUE ;
  Params:=ECachePAR.Items[i] ;
  with Params^ do
    begin
  //  if SurQuelCpI<>zParams.SurQuelCpI then begin Result:=FALSE ; Exit ; end ;
  //  if SurQuelCpO<>zParams.SurQuelCpO then begin Result:=FALSE ; Exit ; end ;
    if Fb1<>zParams.Fb1           then bMatch:=FALSE ;
    if Fb2<>zParams.Fb2           then bMatch:=FALSE ;
    if Date1<>zParams.Date1       then bMatch:=FALSE ;
    if Date2<>zParams.Date2       then bMatch:=FALSE ;
    if Exo<>zParams.Exo           then bMatch:=FALSE ;
    if Etab<>zParams.Etab         then bMatch:=FALSE ;
    if Devise<>zParams.Devise     then bMatch:=FALSE ;
    if SetTyp<>zParams.SetTyp     then bMatch:=FALSE ;
    if DevPivot<>zParams.DevPivot then bMatch:=FALSE ;
    if Multiple<>zParams.Multiple then bMatch:=FALSE ;
    end ;
  if bMatch then Result:=i ;
  end ;
end ;

function TZSpeed.HVerifParams(zParams: ZSParams): Integer ;
var Params: PZSPARAMS ; i: Integer ; bMatch: Boolean ;
begin
Result:=-1 ;
for i:=0 to HCachePAR.Count-1 do
  begin
  bMatch:=TRUE ;
  Params:=HCachePAR.Items[i] ;
  with Params^ do
    begin
  //  if SurQuelCpI<>zParams.SurQuelCpI then begin Result:=FALSE ; Exit ; end ;
  //  if SurQuelCpO<>zParams.SurQuelCpO then begin Result:=FALSE ; Exit ; end ;
    if Fb1<>zParams.Fb1           then bMatch:=FALSE ;
    if Fb2<>zParams.Fb2           then bMatch:=FALSE ;
    if Date1<>zParams.Date1       then bMatch:=FALSE ;
    if Date2<>zParams.Date2       then bMatch:=FALSE ;
    if Exo<>zParams.Exo           then bMatch:=FALSE ;
    if Etab<>zParams.Etab         then bMatch:=FALSE ;
    if Devise<>zParams.Devise     then bMatch:=FALSE ;
    if SetTyp<>zParams.SetTyp     then bMatch:=FALSE ;
    if DevPivot<>zParams.DevPivot then bMatch:=FALSE ;
    if Multiple<>zParams.Multiple then bMatch:=FALSE ;
    end ;
  if bMatch then Result:=i ;
  end ;
end ;

procedure TZSpeed.EGetSQLCumul(iCache: Integer) ;
{$IFNDEF EAGLCLIENT}
var Q: TQuery ;
    stSQL: string ;
    Params: PZSParams ;
    T: TOB ;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
if (iCache<0) or (iCache>ECachePAR.Count) then Exit ;
Params:=ECachePAR.Items[iCache] ;
//Q:=TQuery.Create(Application) ; Q.DataBaseName:=DBSOC.DataBaseName ;
with Params^ do
  begin
  stSQL:=FabricSpeedReq1(fb1, SetTyp, Devise<>'', Etab<>'',
                         Exo<>'', DevPivot, Multiple, Critere,
                         '', Date1, Date2, bAnalyse,Devise,Etab,Exo) ;
  end ;
//Q.SQL.Add(stSQL) ; ChangeSQL(Q) ; PrepareSQLODBC(Q) ;
(*
with Params^ do
  begin
  if Devise<>'' then Q.ParamByName('DEV').AsString:=Devise ;
  if Etab<>''   then Q.ParamByName('ETAB').AsString:=Etab ;
  if Exo<>''    then Q.ParamByName('EXO').AsString:=Exo ;
  end ;
*)
T:=TOB.Create('VCUMUL', nil, -1) ;

//Q.Open ;
Q:=OpenSQL(stSQL, TRUE) ;
if not Q.EOF then
  begin
  while not Q.EOF do begin TOB.CreateDB('VCOMPTE', T, -1, Q) ; Q.Next ; end ;
  end ;
//Q.Close ; Q.Free ;
Ferme(Q) ;
ECacheTOB.Add(T) ;
Params^.bSQLCumul:=TRUE ;
{$ENDIF}
end ;

procedure TZSpeed.HGetSQLCumul(iCache: Integer) ;
{$IFNDEF EAGLCLIENT}
var Q: TQuery ; stSQL, Where: string ; Params: PZSParams ; T: TOB ;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
if (iCache<0) or (iCache>HCachePAR.Count) then Exit ;
Params:=HCachePAR.Items[iCache] ;
Q:=TQuery.Create(Application) ; Q.DataBaseName:=DBSOC.DataBaseName ;
Where:='HB_TYPEBAL="'+Params^.TypePlan+'"'
      +' AND HB_EXERCICE="'+Params^.Exo+'"'
      +' AND HB_DATE1="'+USDateTime(Params^.Date1)+'"'
      +' AND HB_DATE2="'+USDateTime(Params^.Date2)+'"' ;
stSQL:='SELECT HB_COMPTE1, HB_DEBIT AS HB_SUMDEBIT, HB_CREDIT AS HB_SUMCREDIT FROM HISTOBAL WHERE '
      +Where
      +' ORDER BY HB_TYPEBAL, HB_EXERCICE, HB_COMPTE1' ;
Q.SQL.Add(stSQL) ; ChangeSQL(Q) ; PrepareSQLODBC(Q) ;
T:=TOB.Create('VCUMUL', nil, -1) ;
Q.Open ;
if not Q.EOF then
  begin
  while not Q.EOF do begin TOB.CreateDB('VCOMPTE', T, -1, Q) ; Q.Next ; end ;
  end ;
Q.Close ; Q.Free ;
HCacheTOB.Add(T) ;
Params^.bSQLCumul:=TRUE ;
{$ENDIF}
end ;

(*
  Type TabTot = Array[0..7] Of TabDC ;
  0 : AN
  1 : Courantes
  2 : Simu
  3 : Prevision
  4 : Situation
  5 : Révision
  6 : Ce que l'on ne trouve pas
  7 : Total de 1..5 (0 = AN Ne concerne que les écritures normales !!! + Flip Flop 1 <-> 7
E_QUALIFPIECE="N" => Ecritures courantes
E_QUALIFPIECE="S" => Ecriture de simulation
E_QUALIFPIECE="P" => Ecriture de prévision
E_QUALIFPIECE="U" => Ecriture de Situation
E_QUALIFPIECE="R" => Ecriture de révision
E_QUALIFPIECE="C" => Ecriture de clôture
*)
procedure TZSpeed.CalculDC(iCache: Integer; P: string; T: TOB; var lCalc: TabTot; nbDec: Integer; bHisto: Boolean) ;
var TypeEcr, TypeAno: string; D, C : Double ; i: Integer ; (*Stream: TextFile ;*)
begin
(*TODEL
AssignFile(Stream, ExtractFilePath(Application.ExeName)+'CALCUL.TXT');
{$I-} Append(Stream); {$I+}
if IOResult<>0 then Rewrite(Stream) ;
WriteLn(Stream, Format('%s (%s) %8.2f %8.2f',
                      [T.GetValue(P+'GENERAL'),
                       stSolde,
                       Double(T.GetValue(P+'SUMDEBIT')),
                       Double(T.GetValue(P+'SUMCREDIT'))]));
Flush(Stream) ;
CloseFile(Stream) ;
TODEL*)
if bHisto then
  begin
  TypeEcr:='N' ;
  TypeAno:='N' ;
  end else
  begin
  TypeEcr:=T.GetValue(P+'QUALIFPIECE') ;
  TypeAno:=T.GetValue(P+'ECRANOUVEAU') ;
  end ;
if bHisto then HCachePAR.Items[iCache] else ECachePar.Items[iCache] ;
D:=T.GetValue(P+'SUMDEBIT') ;     C:=T.GetValue(P+'SUMCREDIT') ;     
i:=6 ;
if TypeAno<>'N' then i:=0 else
  if TypeEcr='N' then i:=1 else
    if TypeEcr='S' then i:=2 else
      if TypeEcr='P' then i:=3 else
        if TypeEcr='U' then i:=4 else
          if TypeEcr='R' then i:=5 ;
if TypeAno='C' then i:=6 ;
if (D<>0) or (C<>0) then
  begin
  lCalc[i].TotDebit:=lCalc[i].TotDebit+D ;
  lCalc[i].TotCredit:=lCalc[i].TotCredit+C ;
  if i<6 then
    begin
    lCalc[7].TotDebit:=lCalc[7].TotDebit+D ;
    lCalc[7].TotCredit:=lCalc[7].TotCredit+C ;
    end ;
  end ;
end ;

procedure TZSpeed.Calcul(var zCalc: TabTot; lCalc: TabTot; nbDec: Integer; stSolde, Signe: string) ;
var Solde: Double ; i: Integer ; bDebit: Boolean ;
begin
if stSolde='SM' then
  begin
  for i:=0 to 7 do
    begin
    Solde:=Arrondi(lCalc[i].TotDebit-lCalc[i].TotCredit, nbDec) ;
    bDebit:=(Solde>=0) ;
    if bDebit then zCalc[i].TotDebit:=Arrondi(zCalc[i].TotDebit+Solde, nbDec)
              else zCalc[i].TotCredit:=Arrondi(zCalc[i].TotCredit+Solde, nbDec) ;
    end ;
  Exit ;
  end ;
if stSolde='SD' then
  begin
  for i:=0 to 7 do
    begin
    Solde:=Arrondi(lCalc[i].TotDebit-lCalc[i].TotCredit, nbDec) ;
    bDebit:=(Solde>=0) ;
    if bDebit then zCalc[i].TotDebit:=Arrondi(zCalc[i].TotDebit+Solde, nbDec) ;
    end ;
  Exit ;
  end ;
if stSolde='SC' then
  begin
  for i:=0 to 7 do
    begin
    Solde:=Arrondi(lCalc[i].TotDebit-lCalc[i].TotCredit, nbDec) ;
    bDebit:=(Solde>=0) ;
    if not bDebit then zCalc[i].TotCredit:=Arrondi(zCalc[i].TotCredit+Solde, nbDec) ;
    end ;
  Exit ;
  end ;
if stSolde='TD' then
  begin
  for i:=0 to 7 do
    zCalc[i].TotDebit:=Arrondi(zCalc[i].TotDebit+lCalc[i].TotDebit, nbDec) ;
  Exit ;
  end ;
if stSolde='TC' then
  begin
  for i:=0 to 7 do
    zCalc[i].TotCredit:=Arrondi(zCalc[i].TotCredit+lCalc[i].TotCredit, nbDec) ;
  Exit ;
  end ;
end ;

// :    <-> Début:Fin
// ***  <-> Tous
// ---  <-> Non utilisé
// #    <-> Non paramétré
// ,    <-> Séparateur
// &    <-> Ajouté ces comptes
// |    <-> Ajouté ces comptes
// ?    <-> Remplacé par _
// *    <-> Remplacé par %
// Exemple : 001,001,***,#,#,#,#,#,#,#:002,001,***,#,#,#,#,#,#,#,&707,&708,|704
// 3 Tables libre paramétrées (10 max)
//
//§§§§§§§§§§§§§§§§§§§§§§§§
// (SM) <-> Solde mixte
// (SD) <-> Solde débiteur
// (SC) <-> Solde créditeur
// (TD) <-> Total débiteur
// (TC) <-> Total créditeur
//§§§§§§§§§§§§§§§§§§§§§§§§
// RB_SIGNERUB <-> POS
// RB_SIGNERUB <-> NEG
procedure TZSpeed.Parser(iCache: Integer; itParse, P: string; SaufL: TStringList; var zCalc: TabTot ; nbDec: Integer; Signe, Formule: string; bHisto: Boolean) ;
var Master, T: TOB ; i, j, iStart, iEnd, iCol: Integer ; bSauf: Boolean ;
    Params: PZSParams ; Cpt1, Cpt2, stSolde, CptePrec, ColName: string ; lCalc: TabTot ;
begin
if bHisto then
  begin
  Params:=HCachePar.Items[iCache] ;
  if Params^.TypeCumul='BAL' then ColName:='COMPTE1' else ColName:='COMPTE2' ;
  end else
  begin
  Params:=ECachePar.Items[iCache] ;
  case Params^.Fb1 Of
    FbGene : ColName:='GENERAL' ;
    FbAux  : ColName:='AUXILIAIRE';
    fbAxe1..fbAxe5 : ColName:='SECTION' ;
    end ;
  end ;
Cpt1:='' ; Cpt2:='' ; stSolde:='' ;
iStart:=Pos('(', itParse) ; iEnd:=Pos(')', itParse) ;
if (iStart>0) and (iEnd>0) then stSolde:=Copy(itParse, iStart+1, iEnd-(iStart+1)) ;
iCol:=Pos(':', itParse) ;
if iCol>0 then
  begin
  if iStart>0 then Cpt1:=Copy(itParse, 1, iCol-1) ;
  Cpt2:=Copy(itParse, iCol+1, iStart-(iCol+1)) ;
  end else
    if iStart>0 then Cpt1:=Copy(itParse, 1, iStart-1) else
      if (iStart=0) and (itParse<>'') then Cpt1:=itParse ;
if stSolde='' then stSolde:=Formule ;    
if bHisto then Master:=HCacheTOB.Items[iCache]
          else Master:=ECacheTOB.Items[iCache] ;
FillChar(lCalc, Sizeof(lCalc), #0) ; CptePrec:='' ;
for i:=0 to Master.Detail.Count-1 do
  begin
  T:=Master.Detail[i] ;
  if (CptePrec<>'') and (T.GetValue(P+ColName)<>CptePrec) then
    begin
    Calcul(zCalc, lCalc, nbDec, stSolde, Signe) ;
    FillChar(lCalc, Sizeof(lCalc), #0) ;
    end ;
  if Cpt2<>'' then
    begin
    if (Copy(T.GetValue(P+ColName), 1, Length(Cpt1))<Cpt1) or
       (Copy(T.GetValue(P+ColName), 1, Length(Cpt2))>Cpt2) then Continue ;
    end else
      if Copy(T.GetValue(P+ColName), 1, Length(Cpt1))<>Cpt1 then Continue ;
  if SaufL<>nil then
    begin
    bSauf:=FALSE ;
    for j:=0 to SaufL.Count-1 do
      if Copy(T.GetValue(P+ColName), 1, Length(SaufL.Strings[j]))=SaufL.Strings[j]
        then begin bSauf:=TRUE ; Break ; end ;
    if bSauf then Continue ;
    end ;
  CalculDC(iCache, P, T, lCalc, nbDec, bHisto) ;
  CptePrec:=T.GetValue(P+ColName) ;
  end ;
Calcul(zCalc, lCalc, nbDec, stSolde, Signe) ;
end ;

function TZSpeed.GetResult(Total : TabTot; SetTyp : SetttTypePiece; Ano, Signe: string) : Double ;
begin
Result:=0 ;
if tpReel in SetTyp    then Result:=Result+(Total[1].TotDebit-Abs(Total[1].TotCredit)) ;
if tpSim  in SetTyp    then Result:=Result+(Total[2].TotDebit-Abs(Total[2].TotCredit)) ;
if tpPrev in SetTyp    then Result:=Result+(Total[3].TotDebit-Abs(Total[3].TotCredit)) ;
if tpSitu in SetTyp    then Result:=Result+(Total[4].TotDebit-Abs(Total[4].TotCredit)) ;
if tpRevi in SetTyp    then Result:=Result+(Total[5].TotDebit-Abs(Total[5].TotCredit)) ;
if tpCloture in SetTyp then Result:=Result+(Total[6].TotDebit-Abs(Total[6].TotCredit)) ;
case Ano[1] of
  'T': Result:=Total[7].TotDebit-Abs(Total[7].TotCredit) ;
  'A': Result:=Total[0].TotDebit-Abs(Total[0].TotCredit) ;
end ;
if Signe='NEG' then Result:=-Result ;
end ;

function TZSpeed.GetCumul(Cpt, Sauf : string; zParams: ZSParams; NbDec: Integer; Signe, Formule: string): TabTot ;
var SaufL: TStringList ; Calc: TabTot ; P, stParse, itParse: string ;
    iCache: Integer ; bHisto: Boolean ;
begin
FillChar(Calc, Sizeof(Calc), #0) ; bHisto:=FALSE ;
if (zParams.TypePlan<>'') then
  begin
  if HCachePAR=nil then HInitCache ;
  if (HCachePAR.Count>CACHE_MAX) then begin HCleanUp ; HInitCache ; end ;
  iCache:=HVerifParams(zParams) ;
  if iCache<0 then begin iCache:=HSetParams(zParams) ; HGetSQLCumul(iCache) ; end
              else begin if iCache>HCacheTOB.Count-1 then HGetSQLCumul(iCache) ; end ;
  P:='HB_' ; bHisto:=TRUE ;
  end else
  begin
  if ECachePAR=nil then EInitCache ;
  if (ECachePAR.Count>CACHE_MAX) then begin ECleanUp ; EInitCache ; end ;
  iCache:=EVerifParams(zParams) ;
  if iCache<0 then begin iCache:=ESetParams(zParams) ; EGetSQLCumul(iCache) ; end ;
  P:='E_' ; if zParams.Fb1 in [fbAxe1..fbAxe5] then P:='Y_' ;
  end ;
if Sauf<>'' then
  begin
  SaufL:=TStringList.Create ;
  stParse:=Sauf ;
  while stParse<>'' do
    begin
    itParse:=ReadTokenSt(stParse) ;
    if itParse='' then Continue ;
    SaufL.Add(itParse) ;
    end ;
  end else SaufL:=nil ;
stParse:=Cpt ;
while stParse<>'' do
  begin
  itParse:=ReadTokenSt(stParse) ;
  if itParse='' then Continue ;
  Parser(iCache, itParse, P, SaufL, Calc, nbDec, Signe, Formule, bHisto) ;
  end ;
if SaufL<>nil then SaufL.Free ;
Result:=Calc ;
end ;

procedure TZSpeed.GetSQLRub(sCodeRub: string) ;
var Q: TQuery ; stSQL: string ;
begin
if Rub<>nil then begin Rub.Free ; Rub:=nil ; end ;
if sCodeRub<>'' then stSQL:='SELECT * FROM RUBRIQUE WHERE RB_RUBRIQUE LIKE "'+sCodeRub+'%"'
                else stSQL:='SELECT * FROM RUBRIQUE WHERE RB_TYPERUB="GEN" AND RB_NATRUB="CPT"' ;
Q:=OpenSQL(stSQL, TRUE) ;
if not Q.EOF then
  begin
  Rub:=TOB.Create('VRUBS', nil, -1) ;
  while not Q.EOF do begin TOB.CreateDB('VRUB', Rub, -1, Q) ; Q.Next ; end ;
  end ;
Ferme(Q) ;
bSQLRub:=TRUE ;
end ;

function TZSpeed.GetRub(CodeRub: string): TOB ;
var T: TOB ; (*Stream : TextFile ;*)
begin
Result:=nil ;
if not bSQLRub then GetSQLRub('') ;
T:=Rub.FindFirst(['RB_RUBRIQUE'], [CodeRub], FALSE) ;
if T<>nil then Result:=T ;
(*TODEL
if Result<>nil then
  begin
  AssignFile(Stream, ExtractFilePath(Application.ExeName)+'CALCUL.TXT');
  {$I-} Append(Stream); {$I+}
  if IOResult<>0 then Rewrite(Stream) ;
  WriteLn(Stream, Format('%s %s §§§§§§§§§§§§§§§§§§§§§§',
                        [T.GetValue('RB_RUBRIQUE'),
                         T.GetValue('RB_LIBELLE')]));
  Flush(Stream) ;
  CloseFile(Stream) ;
  end ;
TODEL*)
end ;

end.
