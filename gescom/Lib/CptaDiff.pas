unit CptaDiff;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, ParamSoc,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     SysUtils, UtilPGI, EntGC, Classes, HMsgBox, SaisUtil, UtilSais,FactComm,uEntCommun;

Function PassageComptaDifferee ( TOBParam : TOB ) : integer ;

implementation

Var TOBOrig,TOBFinale,TOBGene,TOBRefPiece,TobFusion, TobAcomptes : TOB ;
    sTiersCliDef,sTiersFouDef   : String ;
    sTiersCliPart,sTiersFouPart : String ;
    LesNatsCpta,GeneWhereSQL    : String ;
    DateDeb,DateFin             : TDateTime ;
    DejaPasse,GeneErreur        : integer ;
    DetailEntreprise,DetailParticulier,CompteCentral,DistinguerMPRegle,CentraliserRegleBQE : boolean ;
    RibCliDef, RibFouDef   : string; // JT eQualité 11032
    RibCliPart, RibFouPart : string; // JT eQualité 11032

Type T_CumDateCpta = (cdcJour,cdcSemaine,cdcMois) ;

Const CumDateCpta : T_CumDateCpta = cdcMois ;

Const Err_NoSelection =1 ; Err_NoInsertion = 2 ;

type TOB_Diff = class(TOB)
     private
       function YouzTob(T : TOB) : integer;
     public
       procedure LoadFromBuffer ( Buffer : String );
     end;

function TOB_Diff.YouzTob(T : TOB) : integer;
begin
T.ChangeParent(Self,-1) ;
Result:=0 ;
end;

procedure TOB_Diff.LoadFromBuffer ( Buffer : String ) ;
begin
TOBLoadFromBuffer(Buffer,YouzTob) ;
end;

function RibTiersRegroupement(Auxi : string) : string; // JT eQualité 11032
var Qry : TQuery;
begin
  Result := '';
  if Auxi <> '' then
  begin
    Qry := OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="' + Auxi +'" '+
                   'AND R_PRINCIPAL="X"', True,-1,'',true);
    if not Qry.EOF then
      Result := EncodeRIB(Qry.FindField('R_ETABBQ').AsString, Qry.FindField('R_GUICHET').AsString,
                          Qry.FindField('R_NUMEROCOMPTE').AsString, Qry.FindField('R_CLERIB').AsString,
                          Qry.FindField('R_DOMICILIATION').AsString);
    Ferme(Qry);
  end;
end;

Procedure InitDifferee ( TOBParam : TOB ) ;
Var StCumD : String ;
BEGIN
TOBOrig:=TOB.Create('LESORIGINES',Nil,-1) ;
TOBFinale:=TOB.Create('LESDESTINATIONS',Nil,-1) ;
TOBGene:=TOB.Create('LESGENERAUX',Nil,-1) ;
TobFusion := TOB.Create('LESFUSIONS',nil,-1);
TobAcomptes := TOB.Create('ACOMPTES',nil,-1);
sTiersCliDef:=GetParamSoc('SO_GCCLICPTADIFF') ; sTiersCliPart:=GetParamSoc('SO_GCCLICPTADIFFPART') ;
sTiersFouDef:=GetParamSoc('SO_GCFOUCPTADIFF') ; sTiersFouPart:=GetParamSoc('SO_GCFOUCPTADIFFPART') ;
{Période de regroupement}
StCumD:=TOBParam.GetValue('CUMULDATE') ;
if StCumD='JOU' then CumDateCpta:=cdcJour else
 if StCumD='SEM' then CumDateCpta:=cdcSemaine else
  CumDateCpta:=cdcMois ;
{Fourchette de dates}
DateDeb:=TOBParam.GetValue('DATECPTA') ; DateFin:=TOBParam.GetValue('DATECPTA_') ;
{Déja passé}
if TOBParam.GetValue('DEJAPASSE')='-' then DejaPasse:=0
else if TOBParam.GetValue('DEJAPASSE')='X' then DejaPasse:=1
else if TOBParam.GetValue('DEJAPASSE')='' then DejaPasse:=2 ;
{Natures comptables}
LesNatsCpta:=TOBParam.GetValue('GCD_NATURECOMPTA') ;
{Détails}
DetailEntreprise:=(TOBParam.GetValue('DETAILENTREPRISE')='X') ;
DetailParticulier:=(TOBParam.GetValue('DETAILPARTICULIER')='X') ;
CompteCentral:=(TOBParam.GetValue('COMPTECENTRAL')='X') ;
DistinguerMPRegle:=(TOBParam.GetValue('DISTINGUERMPREGLE')='X') ;
CentraliserRegleBQE:=(TOBParam.GetValue('CENTRALISERREGLEBQE')='X') ;
// JT eQualité 11032
if CompteCentral then
begin
  RibCliDef  := RibTiersRegroupement(sTiersCliDef);
  RibCliPart := RibTiersRegroupement(sTiersCliPart);
  RibFouDef  := RibTiersRegroupement(sTiersFouDef);
  RibFouPart := RibTiersRegroupement(sTiersFouPart);
end else
begin
  RibCliDef := ''; RibFouDef := '';
  RibCliPart := ''; RibFouPart := '';
end;
// Fin JT
END ;

Procedure FinDifferee ;
BEGIN
TOBOrig.Free ;
TOBFinale.Free ;
TOBGene.Free ;
TobFusion.Free;
TobAcomptes.Free;
END ;

procedure ChargerTOBRefPiece(RefPiece:String) ;
Var TOBP: TOB ;
begin
TOBP:=TOB.Create('_Piece',TOBRefPiece,-1) ;
TOBP.AddChampSup('NATURE',False) ;
TOBP.AddChampSup('SOUCHE',False) ;
TOBP.AddChampSup('NUMERO',False) ;
TOBP.AddChampSup('INDICE',False) ;
TOBP.PutValue('NATURE',ReadTokenSt(RefPiece)) ;
TOBP.PutValue('SOUCHE',ReadTokenSt(RefPiece)) ;
ReadTokenSt(RefPiece) ;
TOBP.PutValue('NUMERO',ReadTokenSt(RefPiece)) ;
TOBP.PutValue('INDICE',ReadTokenSt(RefPiece)) ;
end ;

Procedure ChargerOrigines ;
Var Q, QAcptes : TQuery ;
    SQL,Buffer,StNat,RefPiece,RefPiece1,NatureCpta : String ;
    TOBLue : TOB_Diff ;
    TOBEcr, TobTmp : TOB ;
    i : integer ;
    Trouv  : boolean ;
BEGIN
TOBRefPiece:=TOB.Create('_LesPieces',nil,-1) ; // Utilisé pour mettre à jour gp_vivante
SQL:='SELECT * FROM COMPTADIFFEREE WHERE ' ;
GeneWhereSQL:='GCD_DATEPIECE>="'+UsDateTime(DateDeb)+'" AND GCD_DATEPIECE<="'+UsDateTime(DateFin)+'"' ;
Case DejaPasse of
   0 : GeneWhereSQL:=GeneWhereSQL+' AND GCD_COMPTABILISE="-"' ;
   1 : GeneWhereSQL:=GeneWhereSQL+' AND GCD_COMPTABILISE="X"' ;
   2 : ;
   END ;
StNat:=FabricWhereToken(LesNatsCpta,'GCD_NATURECOMPTA') ; if StNat<>'' then GeneWhereSQL:=GeneWhereSQL+' AND '+StNat ;
SQL:=SQL+GeneWhereSQL ;
Q:=OpenSQL(SQL,True,-1,'',true); ; Trouv:=False ;
While Not Q.EOF do
   BEGIN
   RefPiece:=Q.FindField('GCD_REFPIECE').AsString ;
   RefPiece1 := RefPiece;
   NatureCpta := Q.FindField('GCD_NATURECOMPTA').AsString ;
   ChargerTOBRefPiece(RefPiece) ;
   //Tob des acomptes liés
   if (NatureCpta='RC') or (NatureCpta='OC') or (NatureCpta='RF') or (NatureCpta='OF') then
   begin
     TobTmp := TobRefPiece.Detail[TobRefPiece.Detail.count-1];
     QAcptes := OpenSQL('SELECT * FROM ACOMPTES '+
                 'WHERE GAC_NATUREPIECEG = "'+TobTmp.GetValue('NATURE')+'" '+
                 'AND GAC_SOUCHE = "'+TobTmp.GetValue('SOUCHE')+'" '+
                 'AND GAC_NUMERO = "'+TobTmp.GetValue('NUMERO')+'" '+
                 'AND GAC_INDICEG = "'+TobTmp.GetValue('INDICE')+'"',True,-1,'',true);
     TobAcomptes.LoadDetailDB('ACOMPTES', '', '', QAcptes, True);
     Ferme(QAcptes);
   end;
   //Fin Acomptes
   Buffer:=Q.FindField('GCD_BLOBECRITURE').AsString ;
   TOBLue:=TOB_Diff.Create('LESECRITURES',Nil,-1) ;
   TOBLue.LoadFromBuffer(Buffer) ;
   TOBEcr:=TOB.Create('ECRITURES',TOBOrig,-1) ;
   for i:=TOBLue.Detail[0].Detail.Count-1 downto 0 do
   begin
     TobLue.Detail[0].Detail[i].AddChampSupValeur('LANATURECOMPTA',NatureCpta,False);
     TobLue.Detail[0].Detail[i].AddChampSupValeur('LAREFPIECE',RefPiece1,False);
     TOBLue.Detail[0].Detail[i].ChangeParent(TOBEcr,0) ;
   end;
   TOBLue.Free ; Trouv:=True ;
   Q.Next ;
   END ;
Ferme(Q) ;
if Not Trouv then GeneErreur:=Err_NoSelection ;
END ;

Procedure SommerTOBMontants ( TOBE,TOBD : TOB ) ;
BEGIN
AjouteTOBMontant(TOBE,TOBD,'E_DEBIT')     ; AjouteTOBMontant(TOBE,TOBD,'E_CREDIT') ;
AjouteTOBMontant(TOBE,TOBD,'E_DEBITDEV')  ; AjouteTOBMontant(TOBE,TOBD,'E_CREDITDEV') ;
AjouteTOBMontant(TOBE,TOBD,'E_DEBITEURO') ; AjouteTOBMontant(TOBE,TOBD,'E_CREDITEURO') ;
AjouteTOBMontant(TOBE,TOBD,'E_QTE1')      ; AjouteTOBMontant(TOBE,TOBD,'E_QTE2') ;
END ;

Procedure UneSeuleLigneTiers ( NewOrig : TOB ) ;
Var i,k : integer;
    Gene,Auxi,Genetest,AuxiTest : String ;
    TOBEcr,TOBTest : TOB ;
BEGIN
for i:=0 to NewOrig.Detail.Count-1 do
    BEGIN
    TOBEcr:=NewOrig.Detail[i] ;
    Auxi:=TOBEcr.GetValue('E_AUXILIAIRE') ; if Auxi='' then Continue ;
    Gene:=TOBEcr.GetValue('E_GENERAL') ;
    for k:=i+1 to NewOrig.Detail.Count-1 do
        BEGIN
        TOBTest:=NewOrig.Detail[k] ;
        AuxiTest:=TOBTest.GetValue('E_AUXILIAIRE') ; if AuxiTest<>Auxi then Continue ;
        GeneTest:=TOBTest.GetValue('E_GENERAL') ; if GeneTest<>Gene then Continue ;
        SommerTOBMontants(TOBTest,TOBEcr) ;
        TOBTest.PutValue('E_JOURNAL','') ;
        END ;
    END ;
for i:=NewOrig.Detail.Count-1 downto 0 do
    BEGIN
    TOBEcr:=NewOrig.Detail[i] ;
    if TOBEcr.GetValue('E_JOURNAL')='' then TOBEcr.Free ;
    END ;
END ;

Procedure CreatTobFusion(NumFille : integer; TobProv : TOB);
var TobTmp : TOB;
begin
  if (TobProv = nil) or (TobProv.Detail.count = 0) then exit;
  TobTmp := TOB.Create('',TobFusion,-1);
  TobTmp.AddChampSupValeur('FIN_NUM',NumFille,False); //N° fille de la tob finale
  TobTmp.AddChampSupValeur('ORIG_LAREFPIECE',TobProv.Detail[0].GetValue('LAREFPIECE'));
  TobTmp.AddChampSupValeur('ORIG_LANATURECPTA',TobProv.Detail[0].GetValue('LANATURECOMPTA'));
  TobTmp.AddChampSupValeur('ORIG_JALECR',TobProv.Detail[0].GetValue('E_JOURNAL'));
  TobTmp.AddChampSupValeur('ORIG_NUMECR',TobProv.Detail[0].GetValue('E_NUMEROPIECE'));
end;

Procedure CumulerTOB ( NewOrig,TOBDest : TOB );
Var io : integer ;
    Gene,Auxi,TVA,TPF,TVAEnc : String ;
    TOBE,TOBD,TOBIns : TOB ;
    Compatible,Debit : boolean ;
BEGIN
for io:=0 to NewOrig.Detail.Count-1 do
  BEGIN
    TOBE:=NewOrig.Detail[io] ;
    Gene:=TOBE.GetValue('E_GENERAL') ; Auxi:=TOBE.GetValue('E_AUXILIAIRE') ;
    TVA:=TOBE.GetValue('E_TVA') ;      TPF:=TOBE.GetValue('E_TPF') ;
    TVAEnc:=TOBE.GetValue('E_TVAENCAISSEMENT') ;
    TOBD:=TOBDest.FindFirst(['E_GENERAL','E_AUXILIAIRE','E_TVA','E_TPF','E_TVAENCAISEMENT'],[Gene,Auxi,TVA,TPF,TVAEnc],False) ;
    if TOBD<>Nil then
    BEGIN
      Debit:=(TOBE.GetValue('E_DEBIT')<>0) ;
      if Debit then Compatible:=(TOBD.GetValue('E_DEBIT')<>0) else Compatible:=(TOBD.GetValue('E_CREDIT')<>0) ;
    END else Compatible:=False ;
    if Compatible then
    BEGIN
      CreatTobFusion(Tobd.GetIndex, TobE);
      SommerTOBMontants(TOBE,TOBD) ;
    END else
    BEGIN
      TOBIns:=TOB.Create('ECRITURE',TOBDest,-1) ;
      TOBIns.Dupliquer(TOBE,True,True) ;
      CreatTobFusion(TobFinale.detail.count, TobE);
    END ;
  END ;
END ;

Function DetaillerEntreprisePart ( ParentNewOrig : TOB ) : Boolean ;
Var i : integer ;
    TOBE : TOB ;
    PremAux : String ;
BEGIN
Result:=False ; PremAux:='' ;
if ((Not DetailEntreprise) and (Not DetailParticulier)) then Exit ;
if ((DetailEntrePrise) and (DetailParticulier)) then BEGIN Result:=True ; Exit ; END ;
for i:=0 to ParentNewOrig.Detail.Count-1 do
    BEGIN
    TOBE:=ParentNewOrig.Detail[i] ;
    PremAux:=TOBE.GetValue('E_AUXILIAIRE') ;
    if PremAux<>'' then Break ;
    END ;
if PremAux='' then Exit ;
if DetailEntreprise then
   BEGIN
   if ExisteSQL('SELECT T_PARTICULIER FROM TIERS WHERE T_AUXILIAIRE="'+PremAux+'" AND T_PARTICULIER="X"') then Exit ;
   END else
   BEGIN
   if ExisteSQL('SELECT T_PARTICULIER FROM TIERS WHERE T_AUXILIAIRE="'+PremAux+'" AND T_PARTICULIER<>"X"') then Exit ;
   END ;
Result:=True ;
END ;

Function TrouverPiecePourCumul ( ParentNewOrig : TOB ; AuxiPiece, GenePiece, MPPiece, BQEPiece : String ) : TOB ;
Var Jal, NatEcr, CodeD, Regime, Valide, Etab, NCDate : String ;
    TauxDev, Cotation : Double ;
    TOBTrouv, NewOrig,TOBTmp : TOB ;
    AvecAuxi, OkRegle ,AvecBQE, OkCli, DistingEP : Boolean ;
    VCDate : Variant ;
BEGIN
  Result:=Nil ; AvecAuxi:=False ; AvecBQE:=False ;
  if ParentNewOrig.Detail.Count<=0 then Exit ;
  TOBTmp:=ParentNewOrig.Detail[0] ;
  NatEcr := TOBTmp.GetValue('E_NATUREPIECE') ;
  OkRegle := ((NatEcr='RC') or (NatEcr='OC') or (NatEcr='RF') or (NatEcr='OF')) ;
  OkCli := ((NatEcr='FC') or (NatEcr='RC') or (NatEcr='AC') or (NatEcr='OC')) ;
  if OkCli then DistingEP := (sTiersCliDef <> sTiersCliPart) else
  DistingEP := (sTiersFouDef <> sTiersFouPart) ;
  if ((Not CompteCentral) and (BQEPiece<>'') and (OkRegle) and (CentraliserRegleBQE)) then
  begin
    NewOrig:=ParentNewOrig.FindFirst(['E_GENERAL'],[BQEPiece],True) ;
    if NewOrig<>Nil then AvecBQE:=True ;
  end else if ((AuxiPiece<>'') and (Not CompteCentral)) then
  begin
    NewOrig:=ParentNewOrig.FindFirst(['E_AUXILIAIRE'],[AuxiPiece],True) ;
    if NewOrig<>Nil then AvecAuxi:=True ;
  end else if ((CompteCentral) and (DistingEP)) then
  begin
    NewOrig:=ParentNewOrig.FindFirst(['E_AUXILIAIRE'],[AuxiPiece],True) ;
    if NewOrig<>Nil then AvecAuxi:=True ;
  end else
  begin
    NewOrig:=Nil ;
  end ;

  if NewOrig=Nil then
    NewOrig:=ParentNewOrig.Detail[0] ;
  Jal:=NewOrig.GetValue('E_JOURNAL');
  CodeD:=NewOrig.GetValue('E_DEVISE');
  Regime:=NewOrig.GetValue('E_REGIMETVA');
  TauxDev:=NewOrig.GetValue('E_TAUXDEV');
  Cotation:=NewOrig.GetValue('E_COTATION');
  Valide:=NewOrig.GetValue('E_VALIDE');
  Etab:=NewOrig.GetValue('E_ETABLISSEMENT');

  if CumDateCpta=cdcMois then begin NCDate:='E_PERIODE' ; VCDate:=NewOrig.GetValue('E_PERIODE') ; end else
   if CumDateCpta=cdcSemaine then begin NCDate:='E_SEMAINE' ; VCDate:=NewOrig.GetValue('E_SEMAINE') ; end else
    if CumDateCpta=cdcJour then begin NCDate:='E_DATECOMPTABLE' ; VCDate:=NewOrig.GetValue('E_DATECOMPTABLE') ; end ;
  if AvecBQE then
  begin
    {Pour les natures de type règlement, et si option activée, et que l'on ne regroupe pas sur un auxiliaire de centralisation
    alors conservation du détail des "411 Clients" mais cumul en une seule ligne de "512 Banque"}
    if ((MPPiece<>'') and (DistinguerMPRegle)) then
    begin
      TOBTrouv:=TOBFinale.FindFirst(['E_GENERAL','E_JOURNAL','E_NATUREPIECE','E_DEVISE','E_REGIMETVA','E_TAUXDEV','E_COTATION','E_VALIDE','E_ETABLISSEMENT',NCDate,'E_MODEPAIE'],
                                    [BQEPiece,Jal,NatEcr,CodeD,Regime,TauxDev,Cotation,Valide,Etab,VCDate,MPPiece],True) ;
    end else
    begin
      TOBTrouv:=TOBFinale.FindFirst(['E_GENERAL','E_JOURNAL','E_NATUREPIECE','E_DEVISE','E_REGIMETVA','E_TAUXDEV','E_COTATION','E_VALIDE','E_ETABLISSEMENT',NCDate],
                                    [BQEPiece,Jal,NatEcr,CodeD,Regime,TauxDev,Cotation,Valide,Etab,VCDate],True) ;
    end ;
  end else if AvecAuxi then
  begin
    {Cumul des lignes de pièces différentes mais sur auxiliaire identique}
    TOBTrouv:=TOBFinale.FindFirst(['E_AUXILIAIRE','E_GENERAL','E_JOURNAL','E_NATUREPIECE','E_DEVISE','E_REGIMETVA','E_TAUXDEV','E_COTATION','E_VALIDE','E_ETABLISSEMENT',NCDate],
                                  [AuxiPiece,GenePiece,Jal,NatEcr,CodeD,Regime,TauxDev,Cotation,Valide,Etab,VCDate],True) ;
  end else if ((OkRegle) and (MPPiece<>'') and (DistinguerMPRegle)) then
  begin
    {Une écriture regroupée pour les règlements pr mode de paiement}
    TOBTrouv:=TOBFinale.FindFirst(['E_JOURNAL','E_NATUREPIECE','E_DEVISE','E_REGIMETVA','E_TAUXDEV','E_COTATION','E_VALIDE','E_ETABLISSEMENT',NCDate,'E_MODEPAIE'],
                                  [Jal,NatEcr,CodeD,Regime,TauxDev,Cotation,Valide,Etab,VCDate,MPPiece],True) ;
  end else
  begin
    {Cas général, rechercher une écriture compatible}
    TOBTrouv:=TOBFinale.FindFirst(['E_JOURNAL','E_NATUREPIECE','E_DEVISE','E_REGIMETVA','E_TAUXDEV','E_COTATION','E_VALIDE','E_ETABLISSEMENT',NCDate],
                                  [Jal,NatEcr,CodeD,Regime,TauxDev,Cotation,Valide,Etab,VCDate],True) ;
  end ;
  if TOBTrouv<>Nil then
    Result:=TOBTrouv.Parent ;
END ;

Procedure UnSeulTiersDivers ( NewOrig : TOB ; NatEcr : String ; Var LeStAux : String ) ;
Var TOBEcr : TOB ;
    i      : integer ;
    bPart  : boolean ;
    LeAuxi,StAux,StPart : String ;
    LeRib,RibAux,RibPart : string;
    Q                   : TQuery ;
BEGIN
LeAuxi:='' ; bPart:=False ; LeStAux:='' ;
for i:=0 to NewOrig.Detail.Count-1 do
BEGIN
  TOBEcr:=NewOrig.Detail[i] ;
  LeAuxi:=TOBEcr.GetValue('E_AUXILIAIRE') ;
  if LeAuxi<>'' then Break ;
END ;
if LeAuxi='' then Exit ;
if ((NatEcr='OF') or (NatEcr='RF') or (NatEcr='AF') or (NatEcr='FF')) then
BEGIN
  StAux := sTiersFouDef ;   RibAux := RibFouDef;
  StPart := sTiersFouPart ; RibPart := RibFouPart;
END else
BEGIN
  StAux := sTiersCliDef ;   RibAux := RibCliDef;
  StPart := sTiersCliPart ; RibPart := RibCliPart;
END ;
if StAux<>StPart then
BEGIN
  Q:=OpenSQL('SELECT T_PARTICULIER FROM TIERS WHERE T_AUXILIAIRE="'+LeAuxi+'"',True,-1,'',true);
  if Not Q.EOF then bPart:=(Q.Fields[0].AsString='X') ;
  Ferme(Q) ;
END ;

if bPart then
begin
  if StPart<>'' then
    StAux := StPart;
  if RibPart <> '' then
    RibAux := RibPart;
end;

LeStAux := StAux ;
for i:=0 to NewOrig.Detail.Count-1 do
    BEGIN
    TOBEcr:=NewOrig.Detail[i] ;
    if TOBEcr.GetValue('E_AUXILIAIRE')<>'' then // JT eQualité 11032 (Gestion du RIB)
    begin
      TOBEcr.PutValue('E_AUXILIAIRE',StAux) ;
      TOBEcr.PutValue('E_RIB',RibAux) ;
    end;
    if TOBEcr.GetValue('E_CONTREPARTIEAUX')<>'' then TOBEcr.PutValue('E_CONTREPARTIEAUX',StAux) ;
    if ((NatEcr='RC') or (NatEcr='OC') or (NatEcr='RF') or (NatEcr='OF')) then
      TOBEcr.PutValue('E_RIB',RibAux) ;
    END ;
END ;

Procedure SommeAjouteOrigine ( EcrOrig : TOB ) ;
Var TOBDest,NewOrig,TOBEcr : TOB ;
    i               : integer ;
    NatEcr,AuxiPiece,GenePiece,MPPiece,BQEPiece,LeStAux : String ;
BEGIN
if EcrOrig=Nil then Exit ;
if EcrOrig.Detail.Count<=0 then Exit ;
NewOrig:=TOB.Create('',Nil,-1) ; NewOrig.Dupliquer(EcrOrig,True,True) ;
AuxiPiece:='' ; MPPiece:='' ; NatEcr:='' ; LeStAux:='' ;
for i:=0 to NewOrig.Detail.Count-1 do
    BEGIN
    TOBEcr:=NewOrig.Detail[i] ;
    TOBEcr.ClearDetail ; {Virer les analytiques}
    if i=0 then
       BEGIN
       NatEcr:=TOBEcr.GetValue('E_NATUREPIECE') ;
       END ;
    if TOBEcr.GetValue('E_AUXILIAIRE')<>'' then
       BEGIN
       AuxiPiece:=TOBEcr.GetValue('E_AUXILIAIRE') ; GenePiece:=TOBEcr.GetValue('E_GENERAL') ;
       END else
       BEGIN
       BQEPiece:=TOBEcr.GetValue('E_GENERAL') ;
       END ;
    if TOBEcr.GetValue('E_MODEPAIE')<>'' then MPPiece:=TOBEcr.GetValue('E_MODEPAIE') ;
    END ;
if DetaillerEntreprisePart(NewOrig) then
   BEGIN
   TOBDest:=Nil ;
   END else
   BEGIN
   if CompteCentral then UnSeulTiersDivers(NewOrig,NatEcr,LeStAux) ;
   UneSeuleLigneTiers(NewOrig) ; if LeStAux<>'' then AuxiPiece:=LeStAux ; 
   TOBDest:=TrouverPiecePourCumul(NewOrig,AuxiPiece,GenePiece,MPPiece,BQEPiece) ;
   END ;
if TOBDest=Nil then
  BEGIN
    CreatTobFusion(TOBFinale.detail.count,NewOrig);
    NewOrig.ChangeParent(TOBFinale,-1) ;
  END else
  BEGIN
    CumulerTOB(NewOrig,TOBDest) ; NewOrig.Free ;
  END ;
END ;

Procedure AjouteGeneral ( sGene : String ) ;
Var Q : TQuery ;
    TOBG : TOB ;
begin
  TOBG := TOBGene.FindFirst (['G_GENERAL'], [sGene],True) ;
  if TOBG = nil then
  begin
    Q := OpenSQL ('SELECT * FROM GENERAUX WHERE G_GENERAL="' + sGene + '"', True,-1,'',true);
    if Not Q.EOF then TOB.CreateDB ('GENERAUX', TOBGene, -1, Q) ;
    Ferme (Q) ;
  end ;
end ;

Procedure FinirAnalFinale ( TOBEcr : TOB ; NbDec : integer ) ;
Var sGene,Section  : String ;
    TOBG,TOBAxe,TOBAna : TOB ;
    NumAxe : integer ;
    fb     : TFichierBase ;
BEGIN
if TOBEcr=Nil then Exit ;
if TOBEcr.GetValue('E_ANA')<>'X' then Exit ;
sGene:=TOBEcr.GetValue('E_GENERAL') ; AjouteGeneral(sgene) ;
TOBG:=TOBGene.FindFirst(['G_GENERAL'],[sGene],True) ; if TOBG=Nil then Exit ;
for NumAxe:=1 to 5 do if TOBG.GetValue('G_VENTILABLE'+IntToStr(NumAxe))='X' then
    BEGIN
    fb:=AxeToFb('A'+IntToStr(NumAxe)) ;
    TOBAxe:=TOB.Create('A'+InttoStr(NumAxe),TOBEcr,-1) ;
    TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ;
    EcrVersAna(TOBEcr,TOBAna) ;
    Section:=VH^.Cpta[fb].Attente ;
    VentilLigneTOB(TOBAna,Section,1,NbDec,100,TOBEcr.GetValue('E_DEBIT')<>0) ;
    END ;
END ;

Procedure FinirRefLib ( TOBE : TOB ; LaRef : String ) ;
var sGene : String ;
    TOBG  : TOB ;
BEGIN
if TOBE=Nil then Exit ;
sGene:=TOBE.GetValue('E_GENERAL') ; AjouteGeneral(sGene) ;
// JT - eQualité 10168 - Test s'il faut recalculer E_LIBELLE
if (TobE.GetValue('E_AUXILIAIRE') <> '') and ((DetailEntreprise) or (DetailParticulier)) then
begin
  if ExisteSQL('SELECT T_PARTICULIER FROM TIERS WHERE T_AUXILIAIRE="'+TobE.GetValue('E_AUXILIAIRE')+'" AND T_PARTICULIER="X"')then
  begin
    if DetailParticulier then
      Exit;
  end else
  begin
    if DetailEntreprise then
      Exit;
  end;
end;
// Fin JT - eQualité 10168
TOBG:=TOBGene.FindFirst(['G_GENERAL'],[sGene],True) ; if TOBG=Nil then Exit ;
TOBE.PutValue('E_LIBELLE',TOBG.GetValue('G_LIBELLE')) ;
TOBE.PutValue('E_REFINTERNE',LaRef) ;
END ;

Procedure FinirTOBFinale ;
Var i,NewNum,k,Per,Sem,NbDec : integer ;
    Jal,LaRef : String ;
    DD       : TDateTime ;
    DEV      : RDEVISE ;
    TOBEcr,TOBE : TOB ;
    TobTmp, TobTmp1 : TOB;
BEGIN
  for i:=0 to TOBFinale.Detail.Count-1 do
  begin
    TOBEcr:=TOBFinale.Detail[i] ;
    if TOBEcr.Detail.Count<=0 then
      Continue ;
    Jal:=TOBEcr.Detail[0].GetValue('E_JOURNAL') ;
    DEV.Code:=TOBEcr.Detail[0].GetValue('E_DEVISE') ;
    GetInfosDevise(DEV) ;
    NbDec:=DEV.Decimale ;
    Per:=TOBEcr.Detail[0].GetValue('E_PERIODE') ;
    DD:=TOBEcr.Detail[0].GetValue('E_DATECOMPTABLE') ;
    if CumDateCpta=cdcMois then
      DD:=FinDeMois(DD) ;
    Sem:=NumSemaine(DD) ;
    NewNum:=GetNewNumJal(Jal,True,DD) ;
    //JT, maj GAC_NUMECR et GAC_QUALIFPIECE
    TobTmp := TobFusion.FindFirst(['FIN_NUM','ORIG_LANATURECPTA'],[i,TobEcr.Detail[0].GetValue('E_NATUREPIECE')],True);
    while TobTmp <> nil do
    begin
      TobTmp1 := TobAcomptes.FindFirst(['GAC_NUMECR','GAC_JALECR'],[TobTmp.GetValue('ORIG_NUMECR'),TobTmp.GetValue('ORIG_JALECR')],True);
      while TobTmp1 <> nil do
      begin
        TobTmp1.PutValue('GAC_NUMECR',NewNum);
        TobTmp1.PutValue('GAC_QUALIFPIECE','N');
        TobTmp1 := TobAcomptes.FindNext(['GAC_NUMECR','GAC_JALECR'],[TobTmp.GetValue('ORIG_NUMECR'),TobTmp.GetValue('ORIG_JALECR')],True);
      end;
      TobTmp := TobFusion.FindNext(['FIN_NUM','ORIG_LANATURECPTA'],[i,TobEcr.Detail[0].GetValue('E_NATUREPIECE')],True);
    end;
    //Fin JT
    LaRef:=Copy('Centralisation N° '+IntToStr(NewNum)+' / '+DateToStr(DD),1,35) ;
    for k:=0 to TOBEcr.Detail.Count-1 do
    begin
      TOBE:=TOBEcr.Detail[k] ;
      TOBE.PutValue('E_NUMLIGNE',k+1) ;
      TOBE.PutValue('E_QUALIFPIECE','N') ;
      TOBE.PutValue('E_NUMEROPIECE',NewNum) ;
      TOBE.PutValue('E_DATECOMPTABLE',DD) ;
      TOBE.PutValue('E_PERIODE',Per) ;
      TOBE.PutValue('E_SEMAINE',Sem) ;
      if TOBE.GetValue('E_NUMECHE')>0 then
        TOBE.PutValue('E_NUMECHE',1) ;
      FinirAnalFinale(TOBE,NbDec) ;
      FinirRefLib(TOBE,LaRef) ;
    end ;
    MajSoldesEcritureTOB(TOBEcr,True) ;
  end ;
  if Not TOBFinale.InsertDBByNivel(False) then
  begin
    GeneErreur:=Err_NoInsertion
  end else
  begin
    TobAcomptes.UpdateDB(False);
  end;
END ;

Procedure FabriquerDestinations ;
Var io : integer ;
    EcrOrig : TOB ;
BEGIN
  for io:=0 to TOBOrig.Detail.Count-1 do
  begin
    EcrOrig:=TOBOrig.Detail[io] ;
    SommeAjouteOrigine(EcrOrig) ;
  end;
  FinirTOBFinale ;
END ;

Procedure FlagDiff  ;
Var SQL : String ;
BEGIN
SQL:='UPDATE COMPTADIFFEREE SET GCD_COMPTABILISE="X" WHERE GCD_COMPTABILISE="-" AND '+GeneWhereSQL ;
ExecuteSQL(SQL) ;
END ;

Procedure FlagFermeturePiece ;
Var i: Integer ;
TOBP :TOB ;
SQL : String ;
begin
For i:=0 to TOBRefPiece.Detail.count-1 do
  begin
  TOBP:=TOBRefPiece.Detail[i] ;
  if GetInfoParPiece(TOBP.GetValue('NATURE'),'GPP_ACTIONFINI')<>'COM' then continue ; 
  SQL:='UPDATE PIECE SET GP_VIVANTE="-" WHERE GP_NATUREPIECEG="'+TOBP.GetValue('NATURE')+'" AND GP_SOUCHE="'+TOBP.GetValue('SOUCHE')+'" '
                   +' AND GP_NUMERO='+IntToStr(TOBP.GetValue('NUMERO'))
                   +' AND GP_INDICEG='+IntToStr(TOBP.GetValue('INDICE'))+' ';
  ExecuteSQL(SQL) ;
  end ;
TOBRefPiece.Free ;
end ;

Function PassageComptaDifferee ( TOBParam : TOB ) : integer ;
BEGIN
GeneErreur:=0 ;
InitDifferee(TOBParam) ;
ChargerOrigines ;
if GeneErreur=0 then FabriquerDestinations ;
if GeneErreur=0 then FlagDiff ;
if GeneErreur=0 then FlagFermeturePiece ;
FinDifferee ;
Result:=GeneErreur ;
END ;

end.

