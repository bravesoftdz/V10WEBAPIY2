unit ZFolio;

//=======================================================
//=================== Clés primaires ====================
//=======================================================
// ECRITURE : E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE,
//            E_NUMECHE, E_QUALIFPIECE
// ANALYTIQ : Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE,
//            Y_AXE, Y_NUMVENTIL, Y_QUALIFPIECE

interface

uses Classes, 
     {$IFNDEF EAGLCLIENT}
     Db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     {$ENDIF}
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     HCtrls,
     HEnt1,
     HRichOLE,
     UTob,
     Ent1,
     SaisUtil,
     TZ,
     ZCompte,
     ZTypes ,
     ULibEcriture
     ;

type

  TFolio = class
  private
    HistoMontants, HistoDates, HistoDatesAna : TList ;
    FNumFolio         : string ;
    FJal              : RJal ;
    FOpt              : ROpt ;
    TS                : TList ;
    TSA               : TList ;
    TErr              : TStringList ;
    PCompte           : TZCompte ;
    FZTier            : TZTier ;
    FBoRecup          : boolean ;
    FRecupPlantage    : boolean ;
    FStCompteEnErreur : string ;
{$IFDEF SCANGED}
    FListDocId        : TstringList ;
{$ENDIF}
  private
    function    ReadAna(bRestore , bL : Boolean) : Boolean ;
    procedure   DeleteRowNull ; // supprime les lignes sans montant provenant de la saisie
  //  function    EncodeKey : string ;
    procedure   AlimLogSB(St : String) ;
    function    MajComptes : Boolean ;
    function    MajJournal : Boolean ;
    procedure   FaitListeCpt(Clear,Inverse : Boolean) ;
    procedure   CreationLigneAna ( Ecr : TZF );
    function    ReadCompl(bRestore, bL : Boolean) : Boolean ;
  public
    Ecrs            : TOB ;
    EcrsCompl       : TOB ;
    constructor Create(NumFolio: string; Jal: RJal ; Opt: ROpt) ;
    destructor  Destroy ; override ;
    function    CreateRow(Q : TQuery; Row : LongInt) : TZF ;
    procedure   DeleteRow( Row : LongInt ) ;
    function    GetRow(Row : LongInt) : TZF ;
    function    Read(bRestore : Boolean; var TotSavDebit, TotSavCredit : Double; bLibre : Boolean ; bL : boolean = false ; bRecup : boolean = false ) : Boolean ;
    function    Write(Jal : RJal ; vStQualifPiece : string = '' ) : boolean;//;vBoInsertDB : boolean = false) : Boolean ;
    function    Del(Jal : RJal) : Boolean ;
    function    LockFolio : Boolean ;
    function    UnLockFolio ( Par1,Par2,Par3 : string ) : Boolean ; overload ;
    function    UnLockFolio : Boolean ; overload ;
    Procedure   EcritLogSB ;
    function    MajHistos : Boolean ;
    function    MajCumulsCpte(Jal : RJal) : Boolean ;
    procedure   MajCumulsAno ;
    function    GetObmFromRow(Row : LongInt; REdit : THRichEditOLE) : TObm ;
    function    SetObmFromRow(Row : LongInt; REdit : THRichEditOLE ; Obm : TObm) : Boolean ;
    function    GetRowAna(Row, RowAna, RowAxe : LongInt) : TZF ;
    function    CreateRowAna(Q : TQuery; Row, RowAna, RowAxe : LongInt) : TZF ;
    function    IsAnaFromRow(Row : LongInt) : Boolean ;
    procedure   DelAna(Row : LongInt) ;
    procedure   SetAnaOrdre(Row : LongInt; Ordre : Integer) ;
    function    PostAnaCreate(Row : LongInt) : Boolean ;
    procedure   RemuneroteAna ;
    function    RechGuidId( TheTOB : TOB ) : string ;
   {$IFDEF SCANGED}
//   {$IFDEF AGL590}
    procedure   SupprimeLeDocGuid( TheTOB : TOB ) ;
    procedure   AjouteGuidId( TheTOB : TOB ; vGuidId : string);
(*
   {$ELSE}
    procedure   AjouteDocId(vInNum,vInDocId : integer) ;
    procedure   SupprimeDocId ;
    procedure   SupprimeLeDocId ( vInDocId : integer ) ;
    procedure   SupprimeDocIdEcr(vInDocId : integer) ;
    {$ENDIF}
*)
    {$ENDIF} 
    property    Comptes : TZCompte read PCompte write PCompte ;
    property    Tiers   : TZTier   read FZTier   write FZTier ;
    property    RecupPlantage : boolean read FRecupPlantage ;
    property    StCompteEnErreur : string read FStCompteEnErreur ;
  end;

//function  CMajLettrageEnSaisie ( vTOB : TOB ) : boolean ;
procedure ClearLock ;


implementation

uses
{$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  CPVersion,
  CPProcGen,
{$ENDIF MODENT1}
{$IFDEF SCANGED}
  UtilGed ,
{$ENDIF}
  SysUtils,
  hmsgbox,
  ParamSoc,
  ULibAnalytique,
  SaisComm, 
  ed_tools, // pour le VideListe
  UtilPgi;

constructor TFolio.Create(NumFolio : string; Jal : RJal ; Opt : ROpt ) ;
begin
inherited Create;
Ecrs              := TOB.Create('', nil, -1) ;
EcrsCompl         := TOB.Create('', nil, -1) ;
HistoMontants     :=TList.Create ;
HistoDates        :=TList.Create ;
HistoDatesAna     :=TList.Create ;
FNumFolio         :=NumFolio ;
FJal              :=Jal ;
FOpt              :=Opt ;
TS                :=tList.Create ;
TSA               :=TList.Create ;
if NumFolio<>'' then
  Ecrs.Commentaire := DateToStr(EncodeDate(Opt.Year, Opt.Month, 1))+';' +Jal.Code+';'+NumFolio+';' ;
{$IFDEF SCANGED}
FListDocId:=TStringList.Create ;
{$ENDIF}

end ;

destructor TFolio.Destroy ;
var
 i : integer ;
begin
if Ecrs <> nil then Ecrs.Free ;
if EcrsCompl <> nil then EcrsCompl.Free ;
Ecrs          := nil ;
EcrsCompl     := nil ;
VideListe(HistoMontants) ; HistoMontants.Free ;
VideListe(HistoDates) ;    HistoDates.Free ;
VideListe(HistoDatesAna) ; HistoDatesAna.Free ;
VideListe(TS) ; TS.Free ;
for i:=0 to TSA.Count-1 do
 if assigned(TSA[i]) then Dispose(TSA[i]);
TSA.Free ;
{$IFDEF SCANGED}
FListDocId.Free ;
FListDocId :=  nil ;
{$ENDIF}
inherited Destroy ;
end ;

function TFolio.CreateRow(Q : TQuery; Row : LongInt) : TZF ;
var Ecr : TZF ; RowRef : LongInt ;
begin
if Row>Ecrs.Detail.Count-1 then RowRef:=-1
                           else RowRef:=Row ;
if Q<>nil then Ecr:=TZF.CreateDB('ECRITURE', Ecrs, RowRef, Q)
          else Ecr:=TZF.Create('ECRITURE', Ecrs, RowRef) ;
if Q<>nil then Ecr.HistoMontants ;
Result:=Ecr ;
end ;

procedure TFolio.DeleteRow(Row : LongInt) ;
var Ecr : TOB ;
begin
if Row>Ecrs.Detail.Count-1 then Exit ;
Ecr:=Ecrs.Detail[Row] ;
CFreeTOBCompl(Ecr) ;
Ecr.Free ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 16/09/2002
Modifié le ... :   /  /
Description .. : - 16/09/2002 - Supprime les lignes vides provenant de la
Suite ........ : saisie
Mots clefs ... :
*****************************************************************}
procedure TFolio.DeleteRowNull ;
var lEcr : TOB ;
i : integer ;
begin
if (Ecrs=nil) or (Ecrs.DEtail=nil) then exit;
i:=0 ;
while i<Ecrs.Detail.Count do
 begin
  lEcr:=Ecrs.Detail[i] ;
  if (lEcr.GetValue('E_DEBIT')=0) and (lEcr.GetValue('E_CREDIT')=0) then DeleteRow(i)
  else Inc(i) ;
 end; // while
end ;


function TFolio.GetRow(Row : LongInt) : TZF ;
begin
Result:=nil ; if Row>Ecrs.Detail.Count-1 then Exit ;
Result:=TZF(Ecrs.Detail[Row]) ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 08/09/2004
Modifié le ... : 25/03/2005
Description .. : - 08/09/2004 - FB 14531 - blocage si le bordereau contient
Suite ........ : un compte valide
Suite ........ : - LG - 24/03/2005 - FB 15342 - en recup d'ecriture de type
Suite ........ : 'L' on ne test pas la datemodif, on force l'enregsitrement en
Suite ........ : base
Mots clefs ... :
*****************************************************************}
function TFolio.Read(bRestore : Boolean; var TotSavDebit, TotSavCredit : Double; bLibre : Boolean ; bL : boolean = false ; bRecup : boolean = false ) : Boolean ;
var QEcr : TQuery ;
    z : RDel ;
    e : TZF ;
    EcrDebit, EcrCredit : double ;
    lStQualifPiece : string ;
    lStSQL : string ;
    NumCompte : string ;
begin
FBoRecup := bl ;
FRecupPlantage := false ;
//if bL then
lStQualifPiece:=' AND ( E_QUALIFPIECE="N" OR E_QUALIFPIECE="L" ) ' ;
//  else lStQualifPiece:=' AND E_QUALIFPIECE="'+FOpt.QualifP+'" ' ;

Comptes.GetCompteParFolio('E_NUMEROPIECE='+FNumFolio
             +' AND E_JOURNAL="'+FJal.Code+'"'
             +' AND E_EXERCICE="'+FOpt.Exo+'"'
             +' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, 1))+'"'
             +' AND E_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"'
             +lStQualifPiece ) ;

Tiers.Clear ; // 29/05/20007 LG - on remets a zero tout les tiers qd on lit un nouveau bordereau

// SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
QEcr:=OpenSQL('SELECT ' + GetSelectAll('E', True ) + ', E_BLOCNOTE FROM ECRITURE WHERE E_NUMEROPIECE='+FNumFolio
             +' AND E_JOURNAL="'+FJal.Code+'"'
             +' AND E_EXERCICE="'+FOpt.Exo+'"'
             +' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, 1))+'"'
             +' AND E_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"'
             +lStQualifPiece
             +' ORDER BY E_JOURNAL, E_EXERCICE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE', TRUE) ;
Result:=not QEcr.EOF ;
TotSavDebit:=0 ; TotSavCredit:=0;
while not QEcr.EOF do
  begin
  // controle des comptes confidentiels
  NumCompte := QEcr.FindField('E_GENERAL').AsString ;
  if (QEcr.FindField('E_QUALIFPIECE').AsString = 'L') and not FBoRecup then
   begin
    FBoRecup := true ;
    bRecup   := true ;
    bRestore := false ;
    FRecupPlantage := true ;
   end ;
  if (Comptes<>nil) and (Comptes.GetCompte(NumCompte)>-1) and EstConfidentiel(Comptes.GetValue('G_CONFIDENTIEL')) then
   begin
    V_PGI.IOError := oeLettrage ;
    Ferme(QEcr) ;
    exit ;
   end ;
  NumCompte := QEcr.FindField('E_AUXILIAIRE').AsString ;
  if (Tiers<>nil) and (Tiers.GetCompte(NumCompte)>-1) and EstConfidentiel(Tiers.GetValue('T_CONFIDENTIEL')) then
   begin
    V_PGI.IOError := oeLettrage ;
    Ferme(QEcr) ;
    exit ;
   end ;
  // Historique des dates pour la suppression
  z:=RDel.Create ;
  z.NumLigne:=QEcr.FindField('E_NUMLIGNE').AsInteger ;
  z.NumEcheVent:=QEcr.FindField('E_NUMECHE').AsInteger ;
  z.DateModification:=QEcr.FindField('E_DATEMODIF').AsDateTime ;
  z.DateComptable:=QEcr.FindField('E_DATECOMPTABLE').AsDateTime ;
  z.NumFolio:=QEcr.FindField('E_NUMEROPIECE').AsString ;
  z.QualifPiece := QEcr.FindField('E_QUALIFPIECE').AsString ; {JP 20/11/06}
  z.Journal := QEcr.FindField('E_JOURNAL').AsString ;{JP 20/11/06}
  HistoDates.Add(z) ;
  // Historique des montants pour la suppression
  EcrDebit:=QEcr.FindField('E_DEBIT').AsFloat ;
  EcrCredit:=QEcr.FindField('E_CREDIT').AsFloat ;
  if bRestore then
     begin
     TotSavDebit:=TotSavDebit+EcrDebit ;
     TotSavCredit:=TotSavCredit+EcrCredit ;
     end ;
  Ajoute(HistoMontants, 'G', QEcr.FindField('E_GENERAL').AsString,    EcrDebit, EcrCredit,-1,-1,iDate1900) ;
  Ajoute(HistoMontants, 'T', QEcr.FindField('E_AUXILIAIRE').AsString, EcrDebit, EcrCredit,-1,-1,iDate1900) ;
{$IFDEF SCANGED}
(*
{$IFNDEF AGL590}
  if (QEcr.FindField('E_DOCID').asInteger<>0) and (FListDocId.IndexOf(QEcr.FindField('E_DOCID').asString)=-1 )then
   FListDocId.Add(QEcr.FindField('E_DOCID').asString) ;
{$ENDIF}
*)
{$ENDIF}
  if not bRestore then
     begin
        e:=CreateRow(QEcr, QEcr.FindField('E_NUMLIGNE').AsInteger) ;
      //  e.PutValue('SPEEDROW', '-') ;
        e.PutValue('E_CONTROLEUR', '') ;
        E.PutValue('E_QUALIFPIECE', 'N') ;
        if bRecup and ( (e.GetValue('E_ETATLETTRAGE')='TL') or (e.GetValue('E_ETATLETTRAGE')='PL') ) then
        begin
          CSupprimerInfoLettrage(E) ;
          lStSQL := 'UPDATE ECRITURE SET E_DATEPAQUETMAX=E_DATECOMPTABLE,E_IO="X", E_DATEPAQUETMIN=E_DATECOMPTABLE,E_LETTRAGE="",E_ETATLETTRAGE="AL",E_LETTRAGEDEV="-",E_COUVERTURE=0,E_COUVERTUREDEV=0,E_REFLETTRAGE="",E_ETAT="0000000000",';
          lStSQL := lStSQL + 'E_DATEMODIF="' + UsTime(NowH) + '" WHERE E_LETTRAGE="' + e.GetValue('E_LETTRAGE') + '" AND E_GENERAL="' + e.GetValue('E_GENERAL') + '" AND E_AUXILIAIRE="' + e.GetValue('E_AUXILIAIRE') + '"';
          ExecuteSQL(lStSQL);
        end;
     end ;
  QEcr.Next ;
  end ;
Ferme(QEcr) ;
if FRecupPlantage then
 bRestore := false ;
if Result then ReadAna(bRestore,bL) ;
if Result then ReadCompl(bRestore,bL) ;
//Ecrs.SetAllModifie(FALSE) ;
FaitListeCpt(TRUE,TRUE) ;
end ;


Procedure TFolio.FaitListeCpt(Clear,Inverse : Boolean) ;
Var i,j,k,m : integer ;
    Gen,Aux,Sect,Ax : String;
    D,C             : Double ;
    e,ee : TZF ;
    NumP,NumL : Integer ;
    DateP : tDateTime ;
    lInCpt  : Integer; //SG6 27.01.05 Gestion du mode analytique croisaxe
    lIndex : integer ;
    lstNatGene : string ;
BEGIN
lIndex := -1;
If Clear THen TS.Clear ;
for i:=0 to Ecrs.Detail.Count-1 do
    BEGIN
    e:=TZF(Ecrs.Detail[i]) ; //if e.GetValue('BADROW')='X' then Continue ;
    Gen:=e.GetValue('E_GENERAL') ; Aux:=e.GetValue('E_AUXILIAIRE') ;
    D:=e.GetValue('E_DEBIT') ; C:=e.GetValue('E_CREDIT') ;
    If Inverse Then BEGIN D:=D*(-1) ; C:=C*(-1) ; END ;
    NumP:=e.GetValue('E_NUMEROPIECE') ; NumL:=e.GetValue('E_NUMLIGNE') ;
    DateP:=e.GetValue('E_DATECOMPTABLE') ;
    if PCompte <> nil then
     lIndex := PCompte.GetCompte(Gen) ;
    if lIndex > - 1 then
     lStNatGene := PCompte.GetValue('G_NATUREGENE') ;
    AjouteAno(TSA,E,lStNatGene,Inverse) ;
    Ajoute(TS,'G',Gen,D,C,NumP,NumL,DateP) ;
    Ajoute(TS,'T',Aux,D,C,NumP,NumL,DateP,E.GetValue('CPTCEGID'),E.GetValue('YTCTIERS')) ;
    if e.GetValue('E_ANA') = 'X' then
    begin
    CreationLigneAna(e) ;
    for j:=0 to MAXAXE-1 do
        begin
        m:=e.Detail[j].Detail.Count ;
        if m=0 then continue ;
        for k:=0 to m-1 do
            begin
            ee:=TZF(TZF(e.Detail[j]).Detail[k]) ;
            D:=ee.GetValue('Y_DEBIT') ; C:=ee.GetValue('Y_CREDIT') ;
            If Inverse Then BEGIN D:=D*(-1) ; C:=C*(-1) ; END ;

            //SG6 27.01.05 Gestion du mode analytique croisaxe
            if not (VH^.AnaCroisaxe) then
            begin
              Sect:=ee.GetValue('Y_SECTION') ; Ax:=ee.GetValue('Y_AXE') ;
              Ajoute(TS,Ax,Sect,D,C,NumP,NumL,DateP) ;
            end
            else
            begin
              for lInCpt := 1 to MaxAxe do
              begin
                Sect:=ee.GetValue('Y_SOUSPLAN' + IntToStr( lInCpt)) ; Ax:='A' + IntToStr(lInCpt);
                if Sect <> '' then
                begin
                  Ajoute(TS,Ax,Sect,D,C,NumP,NumL,DateP,'','') ;
                end;
              end;
            end;

            end ;
        end ;
    END ;
    end ; // if
END ;

Procedure TFolio.AlimLogSB(St : String) ;
Var St1 : String ;
BEGIN

if not vh^.CPIFDEFCEGID then Exit ;

If TErr=Nil Then TErr:=TStringList.Create ;
Terr.Sorted:=TRUE ; Terr.Duplicates:=DupIgnore ;
St1:='INSERT INTO LOG (LG_UTILISATEUR,LG_MENU,LG_DATE,LG_TEMPS,LG_COMMENTAIRE) VALUES '+
     '("'+V_PGI.User+'",9999,"'+USDATETIME(Date)+'","'+USDATETIME(Date)+'","'+St+'")' ;
TErr.Add(St1) ;
END ;

Procedure TFolio.EcritLogSB ;
Var i : Integer ;
    St : String ;
BEGIN
if not vh^.CPIFDEFCEGID then Exit ;
If TErr=Nil Then Exit ;
For i:=0 To TErr.Count-1 Do BEGIN St:=Terr[i]; ExecuteSQL(St) ; END ;
TErr.Clear ; Terr.Free ; TErr:=Nil ;
END ;

procedure _MajTiersCompl( vTiers, vCompte, vYTCTiers : string ) ;
var
 lTOB : TOB ;
begin
 if (vTiers = '') or (vCompte = '') then exit ;
 lTOB := TOB.Create('TIERSCOMPL',nil,-1) ;
 lTOB.InitValeurs;
 lTOB.PutValue('YTC_AUXILIAIRE',vTiers) ;
 lTOB.PutValue('YTC_TIERS', vYTCTiers) ;
 lTOB.PutValue('YTC_SCHEMAGEN',vCompte) ;
 lTOB.InsertOrUpdateDB ;
 lTOB.Free ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 26/02/2004
Modifié le ... : 09/05/2007
Description .. : - 26/02/2004 - LG - passage en eagl
Suite ........ : - 18/05/2004 - LG - ajout de la mise a jour du compte
Suite ........ : d'acceleration
Suite ........ : - 09/05/2007 - LG - FB 18977 - hcg de section en Axe
Mots clefs ... : 
*****************************************************************}
function TFolio.MajComptes : Boolean ;
var
    i : Integer ; bANouveau, bReseau : Boolean ;
    T : T_SC ;
    StErr : String ;
    FRM : TFRM ;
    lfb : TFichierBase ;
begin
Result:=TRUE ;
bANouveau:=FALSE ; bReseau:=FALSE ;
FaitListeCpt(FALSE,FALSE) ;

for i:=0 to TS.Count-1 do
 begin

  T:=T_SC(TS[i]) ;

  if Arrondi(T.Debit-T.Credit,V_PGI.OkDecV)<>0 then
   begin
    VideTFRM(FRM) ;
    FRM.cpt:=T.Cpte ;
    FRM.axe:=T.Identi ;
    FRM.NumD:=T.NumP ;
    FRM.DateD:=T.DateP ;
    FRM.LigD:=T.NumL ;
    AttribParamsNew ( FRM ,T.Debit,T.Credit , FOpt.TypeExo) ;
    lfb := fbGene ;
    if (T.Identi='T') then
     lfb := fbAux
      else
       if (Copy(T.Identi,1,1)='A') then
        lfb := fbsect ;

    if ExecReqMAJ(lfb,bANouveau,bReseau,FRM) <>1 then
     begin
      if lfb = fbGene then
       FStCompteEnErreur := 'Compte : '
        else
         if lfb = fbAux then
          FStCompteEnErreur := 'Auxiliaire : '
           else
            if lfb = fbSect then
             FStCompteEnErreur := 'Section : ' ;
      FStCompteEnErreur := FStCompteEnErreur + FRM.Cpt ;
      if lfb = fbSect then
       FStCompteEnErreur := FStCompteEnErreur + ' Axe : ' + FRM.Axe ;
      Result:=FALSE ; V_PGI.IoError:=oeSaisie ;
      AlimLogSB(StErr) ;
     end ; // if

   end ; // if

 end ; // for

end ;

function TFolio.MajJournal : boolean ;
var Per : byte ; bANouveau : Boolean ; FRM : TFRM ;
begin
Result:=TRUE ;
bANouveau:=FALSE ;
Fillchar(FRM,SizeOf(FRM),#0) ;
FRM.cpt:=FJal.Code ;
FRM.NumD:=StrToInt(FNumFolio);
FRM.DateD:=FJal.LastDateSais ;
if not bANouveau then
 AttribParamsNew ( FRM ,FJal.TotFolDebit,FJal.TotFolCredit , FOpt.TypeExo)
  else
   begin
    FRM.DE:=FJal.TotFolDebit ;
    FRM.CE:=FJal.TotFolCredit ;
    FRM.DS:=0 ; FRM.CS:=0 ;
    FRM.DP:=0 ; FRM.CP:=0 ;
   end ;

if ExecReqMAJ(fbJal,bANouveau,false,FRM) <>1 then
 begin
  FStCompteEnErreur := 'Journal : ' + FRM.Cpt ;
  Result:=FALSE ;
  V_PGI.IoError:=oeSaisie ;
 end ;

{Dévalidation éventuelle période+jal}
if FOpt.TypeExo=teEncours then
   begin
   Per:=QuellePeriode(FJal.LastDateSais, VH^.Encours.Deb) ;
   if (Per<>0) and (FJal.ValideN[Per]='X') then ADevalider(FJal.Code, FJal.LastDateSais) ;
   end else
   begin
   Per:=QuellePeriode(FJal.LastDateSais, VH^.Suivant.Deb) ;
   if (Per<>0) and (FJal.ValideN1[Per]='X') then ADevalider(FJal.Code, FJal.LastDateSais) ;
   end ;
//SetParamSoc('SO_ZLASTDATE',FJal.LastDateSais) ; SetParamSoc('SO_ZLASTJAL', FJal.Code) ;
end ;


function TFolio.MajCumulsCpte(Jal : RJal) : Boolean ;
begin
{Result:=TRUE ;} FJal:=Jal ;
result := MajComptes and MajJournal ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/06/2006
Modifié le ... :   /  /    
Description .. : - FB 10678 - LG - 02/06/2006 - gestion des devises ds les 
Suite ........ : ano dynamiques
Mots clefs ... : 
*****************************************************************}
procedure TFolio.MajCumulsAno ;
begin
 if not ExecReqMAJAno(TSA) then
  begin
   V_PGI.IoError := oeSaisie ;
   exit ;
  end ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 26/02/2004
Modifié le ... :   /  /
Description .. : - 26/02/2004 - LG - passage en eagl
Mots clefs ... :
*****************************************************************}
function TFolio.MajHistos : Boolean ;
var
  FRM : TFRM ;
begin
Result:=TRUE ;
if (FJal.TotDebDebit<>0) or (FJal.TotDebCredit<>0) then
  begin
   Fillchar(FRM,SizeOf(FRM),#0) ;
   FRM.Cpt := FJal.Code ;
   AttribParamsNew ( FRM ,FJal.TotDebDebit,FJal.TotDebCredit , FOpt.TypeExo) ;
   if ExecReqINVNew(fbJal,FRM) <>1 then begin Result:=FALSE ;V_PGI.IoError:=oeSaisie ; end ;
  end ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 09/04/2002
Modifié le ... : 30/04/2004
Description .. : -Ajout de message en cas d'erreur
Suite ........ : - 18/10/2003 - en renumerotation du bordereau
Suite ........ : ( si des trou sont constate ds le bordereau) insertbynivel ne
Suite ........ : fct pas !
Suite ........ : - 30/04/2004 - LG - recuperation du result de l'insertdb, et
Suite ........ : affectation de V_PGI.IoError en cas de pb
Mots clefs ... :
*****************************************************************}
function TFolio.Write(Jal : RJal ; vStQualifPiece : string = '') : boolean ;//; vBoInsertDB : boolean = false) : Boolean ;
var
 i : integer ;
begin
Result:=TRUE ; FJal:=Jal ;
try
 // pour les ecritures de type L , on ne supprime pas les lignes vides
 if vStQualifPiece <> 'L' then
  DeleteRowNull ;
 CNumeroPiece(Ecrs) ;
 for i := 0 to Ecrs.Detail.Count - 1 do
  begin
   OnDeleteEcritureTOB(Ecrs.Detail[i],taCreat,[cEcrBor]) ;
   OnUpdateEcritureTOB(Ecrs.Detail[i],taCreat,[cEcrBor]) ;
  end ;// for
 Ecrs.SetAllModifie(TRUE) ;
 result:=Ecrs.InsertDBByNivel(true) ;
 if not result then V_PGI.IoError:=oeSaisie ;
 if result then EcrsCompl.InsertDBByNivel(true) ;
except
 On E:Exception do
  begin
   MessageAlerte('Erreur dans l''enregistrement du bordereau !'+#10#13+e.message );
   V_PGI.IoError:=oeSaisie ;
  end;
end ; // except
end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 11/01/2002
Modifié le ... : 25/03/2005
Description .. : LG Modif Affichage d'une message en cas d'erreur de
Suite ........ : suppresion de l'ecriture
Suite ........ : - LG - 24/03/2005 - FB 15342 - en recup d'ecriture de type 
Suite ........ : 'L' on ne test pas la datemodif, on force l'enregsitrement en 
Suite ........ : base
Mots clefs ... : 
*****************************************************************}
function TFolio.Del(Jal : RJal) : boolean ;
var z : RDel ; i : Integer ; Where, WhereHisto : string ; Q : TQuery;
begin
Result:=TRUE ; FJal:=Jal ;
{$IFDEF SCANGED}
(*
{$IFNDEF AGL590}
SupprimeDocId ;
{$ENDIF}
*)
{$ENDIF}
if (HistoDates.Count=0) and (HistoDatesAna.Count=0) then Exit ;
CPStatutDossier ;
for i:=0 to HistoDates.Count-1 do
    begin
    z:=RDel(HistoDates[i]) ;
    Where:='E_JOURNAL="'+FJal.Code+'"'
          +' AND E_EXERCICE="'+FOpt.Exo+'"' ;
    WhereHisto:=' AND E_DATECOMPTABLE="'+UsDateTime(z.DateComptable)+'"'
               +' AND E_NUMLIGNE='+IntToStr(z.NumLigne)
               +' AND E_NUMECHE='+IntToStr(z.NumEcheVent)
               +' AND E_NUMEROPIECE='+z.NumFolio ;
               if not FBoRecup then
                WhereHisto:= WhereHisto + ' AND E_DATEMODIF="'+UsTime(z.DateModification)+'"' ;

    result :=ExecuteSQL('DELETE FROM ECRITURE WHERE '+Where+WhereHisto) > 0 ;
    if FBoRecup then result := true ;

   { if result then
     ExecuteSQL('DELETE FROM ECRCOMPL WHERE '+
                ' EC_JOURNAL="'+FJal.Code+'"'+
                ' AND EC_EXERCICE="'+FOpt.Exo+'"' +
                ' AND EC_DATECOMPTABLE="'+UsDateTime(z.DateComptable)+'"'+
                ' AND EC_NUMLIGNE='+IntToStr(z.NumLigne) +
                ' AND EC_NUMECHE='+IntToStr(z.NumEcheVent) +
                ' AND EC_NUMEROPIECE='+z.NumFolio ) ; }
    if not result then
       BEGIN
        Result:=FALSE ; V_PGI.IOError:=oeSaisie ;
        //LG* affichage d'une message ne cas d'erreur de suppression de l'ecriture
        Q:=OpenSql('select ECRITURE.*,US_LIBELLE from ECRITURE,UTILISAT where ' + Where +
                   ' AND E_DATECOMPTABLE="'+UsDateTime(z.DateComptable)+'"' +
                   ' AND E_NUMLIGNE='+IntToStr(z.NumLigne) +
                   ' AND E_NUMECHE='+IntToStr(z.NumEcheVent) +
                   ' AND E_NUMEROPIECE='+z.NumFolio +
                   ' AND US_UTILISATEUR=E_UTILISATEUR'
                   , true ) ;
        MessageAlerte('La ligne suivante n°' + intToStr(z.NumLigne) + #10#13 +
        DateToStr(z.DateComptable)+' '+ Q.findField('E_LIBELLE').asString+' montant :'+StrS0(Q.findField('E_DEBIT').asFloat)+' '+StrS0(Q.findField('E_CREDIT').asFloat)+
        ' n''a pas pu être mise à jour ! Elle a été modifiée pendant votre saisie par '+Q.findField('US_LIBELLE').asString);
        Ferme(Q) ; Exit ;
       END ;
      {JP 20/11/06 : Branchement des bordereaux en Tréso}
      if EstComptaTreso  and (z.QualifPiece = 'N') then begin
         ExecuteSQL('DELETE FROM ' + GetTableDossier(GetParamSocSecur('SO_TRBASETRESO', ''), 'TRECRITURE') + ' WHERE TE_NODOSSIER = "' + V_PGI.NoDossier + '"' +
                    ' AND TE_NUMTRANSAC = "ICP' + V_PGI.CodeSociete + z.Journal +
                                     Copy(IntToStr(GetPeriode(z.DateComptable)), 3, 4) + z.NumFolio + '" ' +
                    ' AND TE_NUMEROPIECE = ' + z.NumFolio);
      end;
    end ;

 ExecuteSQL('DELETE FROM ECRCOMPL WHERE EC_NUMEROPIECE='+FNumFolio
             +' AND EC_JOURNAL="'+FJal.Code+'"'
             +' AND EC_EXERCICE="'+FOpt.Exo+'"'
             +' AND EC_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, 1))+'"'
             +' AND EC_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"' ) ;


for i:=0 to HistoDatesAna.Count-1 do
    begin
    z:=RDel(HistoDatesAna[i]) ;
    Where:='Y_JOURNAL="'+FJal.Code+'"'
          +' AND Y_EXERCICE="'+FOpt.Exo+'"'
          +' AND Y_NUMEROPIECE='+FNumFolio ;
    WhereHisto:=' AND Y_DATECOMPTABLE="'+UsDateTime(z.DateComptable)+'"'
               +' AND Y_NUMLIGNE='+IntToStr(z.NumLigne)
               +' AND Y_AXE="'+z.Ax+'"'
               +' AND Y_NUMVENTIL='+IntToStr(z.NumEcheVent)
               +' AND Y_DATEMODIF="'+UsTime(z.DateModification)+'"' ;
    result := ( ExecuteSQL('DELETE FROM ANALYTIQ WHERE '+Where+WhereHisto) = 1  ) ;
    if FBoRecup then result := true ;
    if not result then
      BEGIN
        Result:=FALSE ; V_PGI.IOError:=oeSaisie ;
         //LG* affichage d'un message ne cas d'erreur de suppression de l'ecriture
        Q:=OpenSql('select ANALYTIQ.*,US_LIBELLE from ANALYTIQ,UTILISAT where ' + Where +
                   ' AND Y_DATECOMPTABLE="'+UsDateTime(z.DateComptable)+'"' +
                   ' AND Y_NUMLIGNE='+IntToStr(z.NumLigne) +
                   ' AND Y_AXE="'+z.Ax+'"'+
                   ' AND Y_NUMVENTIL='+IntToStr(z.NumEcheVent)+
                   ' AND Y_DATEMODIF="'+UsTime(z.DateModification)+'"'+
                   ' AND US_UTILISATEUR=Y_UTILISATEUR'
                   , true ) ;
        MessageAlerte('La ligne suivante n°' + intToStr(z.NumLigne) + #10#13 +
        DateToStr(z.DateComptable)+' '+ Q.findField('Y_LIBELLE').asString+' montant :'+StrS0(Q.findField('Y_DEBIT').asFloat)+' '+StrS0(Q.findField('Y_CREDIT').asFloat)+
        ' n''a pas pu être mise à jour ! Elle a été modifiée pendant votre saisie par '+Q.findField('US_LIBELLE').asString);
        Ferme(Q) ;
        Exit ;
      END ;
    end ;
{$IFNDEF EAGLCLIENT}
if V_PGI.IOError=oeOk then MajHistos ;
{$ENDIF}
end ;

procedure TFolio.RemuneroteAna ;
var
 i : integer ;
begin
 for i:=0 to Ecrs.Detail.Count -1 do
  Ecrs.Detail[i].Cleardetail ;
 for i:=0 to Ecrs.Detail.Count -1 do
  CreationLigneAna(TZF(Ecrs.Detail[i])) ;
 for i:=0 to Ecrs.Detail.Count -1 do
  begin
   if Ecrs.Detail[i].GetValue('E_ANA')='X' then
    CChargeAna(Ecrs.Detail[i]) ;
  end ;
end ; 


//=======================================================
//================== Gestion des Locks ==================
//=======================================================
{function TFolio.EncodeKey : string ;
begin
Result:=DateToStr(EncodeDate(FOpt.Year, FOpt.Month, 1))+';'+FJal.Code+';'+FNumFolio+';' ;
end ;}

procedure ClearLock ;
begin
BEGINTrans ;
ExecuteSQL('DELETE FROM COURRIER WHERE MG_UTILISATEUR="'+W_W+'" AND MG_TYPE=2000 AND MG_EXPEDITEUR="'+V_PGI.User+'"') ;
CommitTRANS ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 06/11/2002
Modifié le ... :   /  /
Description .. : -06/11/2002- l message d'errer ne s'affichait pas
Mots clefs ... :
*****************************************************************}
function TFolio.LockFolio: Boolean ;
begin
result:=false ;
if FNumFolio<>'' then result:=CBlocageBor(FJal.Code, EncodeDate(FOpt.Year, FOpt.Month, 1) ,StrToInt(FNumFolio),true);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 06/11/2002
Modifié le ... :   /  /
Description .. : -06/11/2002- l message d'errer ne s'affichait pas
Mots clefs ... :
*****************************************************************}
function TFolio.UnLockFolio : Boolean ;
begin
result:=false ;
if FNumFolio<>'' then result:=CBloqueurBor(FJal.Code, EncodeDate(FOpt.Year, FOpt.Month, 1) ,StrToInt(FNumFolio),false,false);
CBloqueurJournal(false,FJal.Code) ;
end;

function TFolio.UnLockFolio ( Par1,Par2,Par3 : string ) : Boolean ;
begin
 result:=CBloqueurBor(Par2, StrToDate(Par1) ,StrToInt(Par3),false);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/07/2002
Modifié le ... :   /  /    
Description .. : - 25/07/2002 - Ajout de la gestion des comps du TOBM
Mots clefs ... : 
*****************************************************************}
function TFolio.GetObmFromRow(Row : LongInt; REdit : THRichEditOLE) : TObm ;
var Ecr : TZF ; 
begin
Result:=nil ; Ecr:=GetRow(Row) ; if Ecr=nil then Exit ;
// Objet 'transfert'
result := TObm.Create(EcrGen, '', TRUE) ;
result.Dupliquer(Ecr,true,true) ;
result.PutValue('E_BLOCNOTE',Ecr.GetValue('E_BLOCNOTE') ) ;
result.CompS:=Ecr.GetValue('COMPS')='X' ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 25/07/2002
Modifié le ... : 08/08/2002
Description .. : - 25/07/2002 - ajout du la gestion de CompS du TOBM ( 
Suite ........ : true si saisie complementaire ]
Suite ........ : 
Suite ........ : - 08/08/2002 - suppression de la destrcution du TOBM -> a 
Suite ........ : faire dasn les code appelant -> plus clair
Mots clefs ... : 
*****************************************************************}
function TFolio.SetObmFromRow(Row : LongInt; REdit : THRichEditOLE ; Obm : TObm) : Boolean ;
var Ecr : TZF ; 
begin
Result:=FALSE ; Ecr:=GetRow(Row) ; if Ecr=nil then begin Obm.Free ; Exit ; end ;
Ecr.Dupliquer(Obm,true,true) ;
If Obm.CompS then Ecr.PutValue('COMPS','X') ;
Result:=TRUE ;
end ;

//=======================================================
//=============== Gestion de l'analytique ===============
//======================================================
{function TFolio.GetStrAxe(Axe : Integer) : string ;
begin
Result:='A'+Chr(48+Axe) ;
end ; }

procedure TFolio.CreationLigneAna ( Ecr : TZF );
var
 i :integer ;
begin
if Ecr.Detail.Count = 0 then
 for i:=1 to MAXAXE do TZF.Create('A'+IntToStr(i), Ecr, -1) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 05/12/2002
Modifié le ... : 13/04/2007
Description .. : - fiche 10478 - la fenetre d'analytique s'ouvrait tous le temps
Suite ........ : si seulement l'axe 2 etait ventilable
Suite ........ : - LG - 13/04/2007 - optimisation de la fct
Mots clefs ... :
*****************************************************************}
function TFolio.IsAnaFromRow(Row : LongInt) : Boolean ;
var Ecr : TZF ;
begin
Result:=FALSE ; Ecr:=GetRow(Row) ; if Ecr=nil then Exit ;
if ( Ecr.Detail=nil ) or (Ecr.Detail.Count = 0)  then Exit ;
//CreationLigneAna(Ecr) ;
result:= (TZF(Ecr.Detail[0]).Detail.Count<>0) or
         (TZF(Ecr.Detail[1]).Detail.Count<>0) or
         (TZF(Ecr.Detail[2]).Detail.Count<>0) or
         (TZF(Ecr.Detail[3]).Detail.Count<>0) or
         (TZF(Ecr.Detail[4]).Detail.Count<>0) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : - LG - 13/04/2007 - optimisation de la fct
Mots clefs ... :
*****************************************************************}
procedure TFolio.DelAna(Row : LongInt) ;
var Ecr : TZF ;
begin
Ecr:=GetRow(Row) ;
if ( Ecr=nil) or ( Ecr.Detail=nil ) or (Ecr.Detail.Count = 0)  then Exit ;
Ecr.ClearDetail ;
//for i:=0 to MAXAXE-1 do Ecr.Detail[i].ClearDetail ;
end ;

procedure TFolio.SetAnaOrdre(Row : LongInt; Ordre : Integer) ;
var Ecr : TZF ; j, k  : Integer ;
begin
Ecr:=GetRow(Row) ; if Ecr=nil then Exit ;
CreationLigneAna(Ecr) ;
for j:=0 to MAXAXE-1 do
  for k:=0 to Ecr.Detail[j].Detail.Count-1 do
    Ecr.Detail[j].Detail[k].PutValue('Y_NUMLIGNE', Ordre) ;
end ;

function TFolio.PostAnaCreate(Row : LongInt) : Boolean ;
var Ecr : TZF ; i, j : Integer ;
begin
Result:=FALSE ; Ecr:=GetRow(Row) ; if Ecr=nil then Exit ;
CreationLigneAna(Ecr) ;
for i:=0 to MAXAXE-1 do
  for j:=0 to TZF(Ecr.Detail[i]).Detail.Count-1 do
    if not TZF(TZF(Ecr.Detail[i]).Detail[j]).FieldExists('OLDDEBIT') then
      begin
      TZF(TZF(Ecr.Detail[i]).Detail[j]).ChampSupAna ;
      TZF(TZF(Ecr.Detail[i]).Detail[j]).HistoMontants ;
      end ;
Result:=TRUE ;
end ;


function TFolio.GetRowAna(Row, RowAna, RowAxe : LongInt) : TZF ;
var Ecr : TZF ;
begin
Result:=nil ;
Ecr:=GetRow(Row) ; if Ecr=nil then Exit ;
if ( Ecr.Detail=nil ) or (Ecr.Detail.Count = 0)  then Exit ;
//CreationLigneAna(Ecr) ;
if RowAna>TZF(Ecr.Detail[RowAxe]).Detail.Count-1 then Exit ;
Result:=TZF(TZF(Ecr.Detail[RowAxe]).Detail[RowAna]) ;
end ;

function TFolio.CreateRowAna(Q : TQuery; Row, RowAna, RowAxe : LongInt) : TZF ;
var Ecr, Ana : TZF ;
begin
Ecr:=GetRow(Row) ; if Ecr=nil then begin Result:=nil; Exit ; end ;
CreationLigneAna(Ecr) ;
if RowAna>TZF(Ecr.Detail[RowAxe]).Detail.Count-1 then
   begin
   if Q<>nil then Ana:=TZF.CreateDB('ANALYTIQ', TZF(Ecr.Detail[RowAxe]), -1, Q)
             else Ana:=TZF.Create('ANALYTIQ', TZF(Ecr.Detail[RowAxe]), -1) ;
   end else
   begin
   if Q<>nil then Ana:=TZF.CreateDB('ANALYTIQ', TZF(Ecr.Detail[RowAxe]), RowAna, Q)
             else Ana:=TZF.Create('ANALYTIQ', TZF(Ecr.Detail[RowAxe]), RowAna) ;
   end ;
if Q<>nil then Ana.HistoMontants ;
Result:=Ana ;
end ;

function TFolio.ReadAna(bRestore, bL : Boolean) : Boolean ;
var QAna : TQuery ; z : RDel ;
    AnaDebit, AnaCredit : double ; AnaSection : string ;
    lStSql : string ;
    lAna : TZF ;
begin
// SBO 05/07/2007 : gestion pb DB2 avec champ blocnote à positionner à la fin
lStSql := 'SELECT ' + GetSelectAll('Y', True ) + ', Y_BLOCNOTE FROM ANALYTIQ WHERE Y_NUMEROPIECE='+FNumFolio
             +' AND Y_JOURNAL="'+FJal.Code+'"'
             +' AND Y_EXERCICE="'+FOpt.Exo+'"'
             +' AND Y_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, 1))+'"'
             +' AND Y_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"' ;
if bL then
 lStSql := lStSql + ' AND ( Y_QUALIFPIECE="'+FOpt.QualifP+'" OR Y_QUALIFPIECE="L" ) '
  else
   lStSql := lStSql + ' AND Y_QUALIFPIECE="'+FOpt.QualifP+'" ' ;
lStSql  := lStSql + ' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL' ;
QAna:=OpenSQL( lStSql, TRUE) ;
Result:=not QAna.EOF ;
while not QAna.EOF do
  begin
  // Historique des dates pour la suppression
  z:=RDel.Create ;
  z.NumLigne:=QAna.FindField('Y_NUMLIGNE').AsInteger ;
  z.NumEcheVent:=QAna.FindField('Y_NUMVENTIL').AsInteger ;
  z.DateModification:=QAna.FindField('Y_DATEMODIF').AsDateTime ;
  z.Ax:=QAna.FindField('Y_AXE').AsString ;
  z.DateComptable:=QAna.FindField('Y_DATECOMPTABLE').AsDateTime ;
  z.QualifPiece := QAna.FindField('Y_QUALIFPIECE').AsString ; {JP 20/11/06}
  z.Journal := QAna.FindField('Y_JOURNAL').AsString ; {JP 20/11/06}
  HistoDatesAna.Add(z) ;
  // Historique des montants pour la suppression
  AnaSection:=QAna.FindField('Y_SECTION').AsString ;
  AnaDebit:=QAna.FindField('Y_DEBIT').AsFloat ;
  AnaCredit:=QAna.FindField('Y_CREDIT').AsFloat ;
  Ajoute(HistoMontants, z.Ax, AnaSection, AnaDebit, AnaCredit,-1,-1,iDate1900) ;
  // Charger le TZF
  if not bRestore then
   begin
    lAna := CreateRowAna(QAna, z.NumLigne-1, z.NumEcheVent-1, CGetNumAxe(z.Ax)-1) ;
    if bL then
     lAna.PutValue('Y_QUALIFPIECE','N' ) ;
   end ;// if
  QAna.Next ;
  end ;
Ferme(QAna) ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 08/06/2005
Modifié le ... :   /  /    
Description .. : - FB 15876 - LG - 08/06/2005 - Lorsque j'indique 
Suite ........ : 01/01/1900, il ne faut pas mettre le triangle bleu sur le jour 
Suite ........ : dans l'écriture
Mots clefs ... : 
*****************************************************************}
function TFolio.ReadCompl(bRestore, bL : Boolean) : Boolean ;
var
 lQ       : TQuery ;
 lStSql   : string ;
 lInIndex : integer ;
 lTOB     : TOB ;
begin
lStSql := 'SELECT * FROM ECRCOMPL WHERE EC_NUMEROPIECE='+FNumFolio
             +' AND EC_JOURNAL="'+FJal.Code+'"'
             +' AND EC_EXERCICE="'+FOpt.Exo+'"'
             +' AND EC_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, 1))+'"'
             +' AND EC_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"' ;
if bL then
 lStSql := lStSql + ' AND ( EC_QUALIFPIECE="'+FOpt.QualifP+'" OR EC_QUALIFPIECE="L" ) '
  else
   lStSql := lStSql + ' AND EC_QUALIFPIECE="'+FOpt.QualifP+'" ' ;
lStSql  := lStSql + ' ORDER BY EC_JOURNAL, EC_EXERCICE, EC_NUMEROPIECE, EC_NUMLIGNE' ;
lQ:=OpenSQL( lStSql,true) ;
Result:=not lQ.EOF ;
while not lQ.EOF do
 begin
  lInIndex := lQ.FindField('EC_NUMLIGNE').AsInteger ;
  if lInIndex <= Ecrs.Detail.Count then
   begin
    lTOB := Ecrs.Detail[lInIndex-1] ;
    lTOB.PutValue('COMPTAG','X') ;
    if lQ.FindField('EC_CUTOFFDEB').asDateTime <> iDate1900 then
     lTOB.PutValue('COMPTAG','XD') ;
    if length(lQ.FindField('EC_DOCGUID').asString) > 1 then
     lTOB.PutValue('COMPTAG',lTOB.GetValue('COMPTAG')+ 'A') ;
    CCreateDBTOBCompl(lTOB,EcrsCompl, lQ) ;
   end ; // if
  lQ.Next ;
 end ; // while
Ferme(lQ) ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 08/10/2002
Modifié le ... :   /  /
Description .. : -08/10/2002 - ajout de l'enregistrement dans la table log
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLCLIENT}
{function TFolio.Restore : Boolean ;
var
 s : string ;
begin
 result:= inherited Restore;
 if result then
  begin
   s:=copy('R '+EncodeKey,1,35) ; // la champs fait 35 caracteres
   executeSQL('INSERT INTO LOG (LG_UTILISATEUR,LG_MENU,LG_DATE,LG_TEMPS,LG_COMMENTAIRE) VALUES '+
              '("'+V_PGI.User+'",7005,"'+UsTime(Date)+'","'+USTIME(Date)+'","'+s+'")' ) ;

  end;
end;}
{$ENDIF}

function CMajLettrageEnSaisie ( vTOB : TOB ) : boolean ;
var
 lStSQL : String ;
begin

 lStSQL:='UPDATE ECRITURE SET E_DATEMODIF="'+UsTime(NowH)+'" ,E_QUALIFPIECE="N",E_IO="X" '
      +' Where E_GENERAL="'+vTOB.GetValue('E_GENERAL')+'" AND E_AUXILIAIRE="'+vTOB.GetValue('E_AUXILIAIRE')+'"'
      +' AND E_DATECOMPTABLE="'+UsDateTime(vTOB.GetValue('E_DATECOMPTABLE'))+'" AND E_NUMEROPIECE='+InttoStr(vTOB.GetValue('E_NUMEROPIECE'))
      +' AND E_QUALIFPIECE="'+vTOB.GetValue('E_QUALIFPIECE')+'" AND E_NATUREPIECE="'+vTOB.GetValue('E_NATUREPIECE')+'"'
      +' AND E_NUMLIGNE='+IntToStr(vTOB.GetValue('E_NUMLIGNE'))+' AND E_NUMECHE='+IntToStr(vTOB.GetValue('E_NUMECHE'))
      +' AND E_DATEMODIF="'+UsTime(vTOB.GetValue('E_DATEMODIF'))+'"' ;
 try
  result:=(ExecuteSQL(lStSQL)=1) ;
  if not result then
     PGIInfo('La ligne suivante : '+#10#13+
     'Exercice : '+vTOB.GetValue('E_EXERCICE')+' journal : '+ vTOB.GetValue('E_JOURNAL')+#10#13+
     'Bordereau n° '+intToStr(vTOB.GetValue('E_NUMEROPIECE'))+' du '+ DateToStr(vTOB.GetValue('E_DATECOMPTABLE'))+' libellé '+ vTOB.GetValue('E_LIBELLE')+
     ' débit : '+StrS0(vTOB.GetValue('E_DEBIT'))+' crédit : '+StrS0(vTOB.GetValue('E_CREDIT'))+#10#13+' n''a pas pu être mise à jour ! ','Erreur dans l''écriture');
 except
  on E:Exception do
   begin
    result:= false ; MessageAlerte(E.Message) ;
   end;
 end;

END ;


function TFolio.RechGuidId( TheTOB : TOB ) : string ;
var
 lTOBEcr   : TOB ;
 lInNum    : integer ;
begin

 result := '' ;

 lInNum  := TheTOB.GetValue('E_NUMGROUPEECR') ;
 lTOBEcr := Ecrs.FindFirst(['E_NUMGROUPEECR'],[lInNum],false) ;

 if lTOBEcr <> nil then
  result := VarAsType(CGetValueTOBCompl(lTOBEcr,'EC_DOCGUID'),varString) ;

 if result = #0 then result := '' ;

end ;

{$IFDEF SCANGED}
//{$IFDEF AGL590}

procedure TFolio.AjouteGuidId( TheTOB : TOB ; vGuidId : string ) ;
var
 lInNum    : integer ;
 lTOBCompl : TOB ;
 lTOBEcr   : TOB ;
begin

 lInNum  := TheTOB.GetValue('E_NUMGROUPEECR') ;
 lTOBEcr := Ecrs.FindFirst(['E_NUMGROUPEECR'],[lInNum],false) ;

 if lTOBEcr = nil then exit ;

 lTOBCompl := CGetTOBCompl(lTOBEcr) ;
 if lTOBCompl = nil then
  lTOBCompl := CCreateTOBCompl(lTOBEcr,EcrsCompl)
   else
    if Length(lTOBCompl.GetValue('EC_DOCGUID')) > 1 then
     SupprimeDocumentGed(lTOBCompl.GetValue('EC_DOCGUID')) ;

 lTOBCompl.PutValue('EC_DOCGUID', vGuidId) ;

end ;



procedure TFolio.SupprimeLeDocGuid( TheTOB : TOB ) ;
var
 lInNum  : integer ;
 lStGuid : string ;
 lTOBEcr : TOB ;
begin           

 lStGuid := RechGuidId(TheTOB) ;             

 if lStGuid = '' then exit ;

 SupprimeDocumentGed(lStGuid) ;

 lInNum  := TheTOB.GetValue('E_NUMGROUPEECR') ;
 lTOBEcr := Ecrs.FindFirst(['E_NUMGROUPEECR'],[lInNum],false) ;
 if lTOBEcr = nil then exit ;
 CPutValueTOBCompl(lTOBEcr,'EC_DOCGUID','') ;
  
end ;
(*
{$ELSE}
procedure TFolio.AjouteDocId(vInNum,vInDocId : integer);
var
 Ecr : TOB ;
 lIndex : integer ;
begin
 Ecr:=Ecrs.FindFirst(['E_NUMGROUPEECR'],[vInNum],false) ; if Ecr=nil then exit ;
 if Ecr.GetValue('E_DOCID')>0 then
  begin
   SupprimeDocumentGed(Ecr.GetValue('E_DOCID')) ;
   lIndex:=FListDocId.IndexOf(intToStr(vInDocId)) ;
   if lIndex>-1 then FListDocId.Delete(lIndex) ;
  end ;
 while Ecr<>nil do
  begin
   Ecr.PutValue('E_DOCID',vInDocId) ;
   Ecr:=Ecrs.FindNext(['E_NUMGROUPEECR'],[vInNum],false) ;
  end ;// while
 if FListDocId.IndexOf(intToStr(vInDocId))=-1 then FListDocId.add(intToStr(vInDocId)) ;
end;

procedure TFolio.SupprimeDocIdEcr(vInDocId : integer);
var
 Ecr : TOB ;
 lIndex : integer ;
begin
 Ecr:=Ecrs.FindFirst(['E_DOCID'],[vInDocId],false) ; if Ecr=nil then exit ;
 while Ecr<>nil do
  begin
   Ecr.PutValue('E_DOCID',0) ;
   Ecr:=Ecrs.FindNext(['E_DOCID'],[vInDocId],false) ;
  end ;// while
 lIndex:=FListDocId.IndexOf(intToStr(vInDocId)) ;
 if lIndex>-1 then FListDocId.Delete(lIndex) ;
end;

procedure TFolio.SupprimeDocId ;
var
 i : integer ;
 lIndex : integer ;
begin
if Ecrs.Detail.Count=0 then exit ;
for i:=0 to Ecrs.Detail.Count-1 do
 begin
  lIndex:=FListDocId.IndexOf(intToStr(Ecrs.Detail[i].GetValue('E_DOCID'))) ;
  if lIndex>-1 then FListDocId.delete(lIndex) ;
 end; // for
for i:=0 to FListDocId.Count - 1 do SupprimeDocumentGed(strToInt(FListDocId[i])) ;
end;


procedure TFolio.SupprimeLeDocId( vInDocId : integer ) ;
begin
 SupprimeDocumentGed(vInDocId) ;
 if ( FListDocId =  nil ) or ( FListDocId.Count = 0 ) then exit ;
 if FListDocId.IndexOf(intToStr(vInDocId))>-1 then
  FListDocId.delete(FListDocId.IndexOf(intToStr(vInDocId))) ;
end;

{$ENDIF}
*)
{$ENDIF}




end.



