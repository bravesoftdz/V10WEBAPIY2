unit SaisComm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls,
  {$IFNDEF EAGLSERVER}
  Forms, Dialogs,
  StdCtrls, Grids,Mask,Buttons,Menus,HPanel,
  {$ENDIF}
  Hctrls,
  ExtCtrls,
  ComCtrls,
  //RRO Hspliter,
  {$IFNDEF SANSCOMPTA}
  {$IFNDEF PGIIMMO}
  HCompte,
  {$ENDIF}
  {$ENDIF}
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB,
{$ENDIF}
{$IFDEF VER150}
  variants,
{$ENDIF}
  HEnt1, UTOB, Paramsoc,
{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
  TimpFic,
{$ENDIF}
{$ENDIF}
  HQry,SaisUtil, ed_tools,UtilPGI
    ,ENT1
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
,uEntCommun
  ;

type TTypeExo = (tePrecedent,teSuivant,teEncours) ;
type TSorteMontant = (tsmPivot,tsmDevise,tsmEuro) ;
Type T_DPE = (tdDevise,tdPivot,tdEuro) ;
Type T_YATP = (yaRien,yaNL,yaL) ;
Type TFRM = Record
            Cpt,Axe : String ;
            Deb,Cre,DE,CE,DS,CS,DP,CP,DebDev,CreDev : Double ;
            DateD : tDateTime ;
            NumD,LigD , NumClo,NumAno : Integer ;
            G_L,G_C,G_N,T_L,T_N : String ;
            G_V : Array[1..MaxAxe] Of String ;
            fb : TFichierBase; // LG* 13/11/2001
            CptCegid : string ;
            Dossier  : String ;
            NatGene  : string ;
            Etabliss : string ;
            Dev      : string ;
            Exo      : string ;
            End ;

Type T_SC = Class
            Identi : String3 ; { 'G', 'T', 'A1'..'A5', 'GS', 'B' }
            Cpte   : String17 ;
            Debit  : Double ;
            Credit : Double ;
            NumP,NumL   : Integer ;
            DateP : tDateTime ;
            CptCegid : string ; // LG 18/04/2004
            YTCTiers : string ;
            end ;
type DelInfo = Class
               LeCod,LeLib,LeMess, LeMess2, LeMess3, LeMess4, Gen, Aux, Sect : String ;
               GenACreer,AuxACreer,SectACreer : Boolean ;
               Jal,Qualif,Axe : String ;
               NumP : Integer ;
               Chrono : Integer ;
               End ;


{$IFNDEF EAGLSERVER}
Function  GetO ( GS : THGrid ; Lig : integer = - 1 ) : TOBM ;
procedure SetO ( GS : ThGrid ; O : TOB ; Lig : integer = -1 ) ;
{$ENDIF}
function  OBMToIdent( O : TOBM ; Totale : boolean ) : RMVT ;

{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
type TSAJAL = Class
              JOURNAL        : String ;
              LIBELLEJOURNAL : String ;
              ABREGEJOURNAL  : String ;
              NATUREJAL,AXE,ModeSaisieJal  : String3 ;
              COMPTEURNORMAL,COMPTEURSIMUL : String3 ;
              ValideEN,ValideEN1           : String[24] ;
              MULTIDEVISE,MontantNegatif : Boolean ;
              COMPTEINTERDIT, COMPTEAUTOMAT : String ;
              OLDDEBIT,OLDCREDIT : Double ;
              NbAuto         : integer ;
              TRESO          : TLigTreso ;
              SAISIETRESO,OkFerme,Effet,JalEffet : Boolean ;
              J_ACCELERATEUR : string ;
             Private
              {JP 17/01/08 : FQ 18563 : Il faut gérer la date et l'exo pour le calcul des soldes}
              FDateRef : TDateTime;
              FTypeExo : TTypeExo;
              procedure SetDateExo(Value : TDateTime);

              function  SelectJAL ( Jal : String ) : String ;
              procedure QueryToJal ( JQ : TQuery ) ;
             Public
              Function   ChargeCompteTreso : Boolean ;
              Constructor Create ( Jal : String ; SaisTreso : Boolean) ;
              Constructor CreateEff ( Jal : String ) ;
              {JP 17/01/08 : FQ 18563 : Il faut gérer la date et l'exo pour le calcul des soldes}
              procedure PutDate(DateC : string);
              property DateRef : TDateTime read FDateRef write SetDateExo;
             END ;

{$IFNDEF EAGLSERVER}
procedure PositionneTz( Cache : THCpteEdit ; ip : integer ) ;
Function  isBQE ( GS : THGrid ; Lig : integer ) : boolean ;
Function  isTiers ( GS : THGrid ; Lig : integer ) : boolean ;
Function  isHT ( GS : THGrid ; Lig : integer ; OkTVA : boolean ) : boolean ;
Function  TrouveLigCompte ( GS : THGrid ; Col,FromLig : integer ; Cpte : String17 ) : integer ;
Function  TrouveLigHT ( GS : THGrid ; FromLig : integer ; OkTVA : boolean ) : integer ;
Function  TrouveLigTiers ( GS : THGrid ; FromLig : integer) : integer ;
Function  TrouveLigTVA ( GS : THGrid ; FromLig : integer ) : integer ;
Function  LaMeme ( GS : THGrid ; L : integer) : Boolean ;
Function  LeMeme ( GS : THGrid ; C,L : integer) : Boolean ;
Function  TrouveSens ( GS : THGrid ; Nat : String3 ; Lig : integer ; CommeBQE : boolean = False ) : byte ;
function  ArretZone( Col : integer ; Var St : String ; MAJ : boolean ) : boolean ;
procedure Desalloue( GS : THGRid ; LaCol,LaLig : Longint ) ;
Procedure AttribGen ( GS : THGrid ; LigC,Lig : integer ) ;
Procedure AttribAux ( GS : THGrid ; LigC,Lig : integer ) ;
{$ENDIF}


procedure AttribParamsComp ( Var FRM : TFRM ; OO : TObject ) ;
{$IFNDEF EAGLSERVER}
Function  GetChamp( GS : THGrid ; Lig : integer ; Quoi : String ) : Variant ;
Function  GetDouble( GS : THGrid ; Lig : integer ; Quoi : String ) : Double ;
{$ENDIF}
//deplacement de la fct Function  GetO ( GS : THGrid ; Lig : integer ) : TOBM ;
{$IFNDEF EAGLSERVER}
procedure NumeroteLignes ( GS : THGrid ) ;
procedure AlloueMvt ( GS : THGrid ; Ecr : TTypeEcr ; Lig : integer ; Init : boolean ) ;
procedure SwapOutil( Outils : TPanel ) ;
function  MontantLigne ( GS : THGrid ; Lig : integer ; tm : TSorteMontant ) : double ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$IFNDEF EAGLSERVER}
Function  ValD ( GS : THGrid ; Lig : integer ) : double ;
Function  ValC ( GS : THGrid ; Lig : integer ) : double ;
Function  EstRempli( GS : THGrid ; Lig : integer) : Boolean ;
procedure BlocageEntete ( F : TForm ; Treso : Boolean ) ;
procedure FormatMontant ( GS : THGrid ; ACol,ARow,Decim : integer ) ;
procedure NewLigne ( GS : THGrid )  ;
procedure VideZone ( GS : THGrid ) ;
{$ENDIF}
Function  EcheMvt ( T : TDataSet ) : Boolean ;
Function  AnalMvt ( T : TDataSet ) : Boolean ;
Function  WhereEcriture ( ts : TTypeSais ; M : RMVT ; Totale : boolean ; ChargeUnBordereau : Boolean = FALSE) : String ;
function  TOBToIdent( O : TOB ; Totale : boolean ) : RMVT ;

function  ExecReqMAJ ( fb : TFichierBase ; Ano,AvecReseau : boolean ; Var FRM : TFRM ) : LongInt ;
function  ExecReqMAJAno( TheList : TList ) : boolean ;
function  FabricReqComp ( fb : TFichierBase  ; Var FRM : TFRM ) : String ;

function  EncodeLC ( R : RMVT ) : String ;
function  DecodeLC ( StM : String ) : RMVT ;
{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}

Function  ExecReqINVNew ( fb : TFichierBase ; Var FRM : TFRM) : LongInt ;

//LG* function  OBMToIdent( O : TOBM ; Totale : boolean ) : RMVT ;
function  OBMToIdentManuel( O : TOBM ; Totale : boolean ; NatPiece : String3) : RMVT ;
Function  CompleteReseau ( ts : TTypeSais ; O : TOBM ) : String ;
procedure ProrateQteAnalObj ( OBA : TObjAna ; O : TOBM ; OQ,NQ : Double ; Ind : integer ) ;
procedure ProrateQteAnalTOB ( vTOB : TOB ;  vRdQte , vRdNQ : Double ; vStInd : string ) ;
{$IFNDEF EAGLSERVER}
procedure ComCom1 ( G : THGrid ; O : TOB ) ;
procedure ComComLigne ( G : THGrid ; O : TOB ; Lig : integer ) ;
{$ENDIF}
function  Scenario2Comp ( Scenario : TOB ) : String ;
Function  ModifRIBOBM (O : TOB ; Maj,Treso : boolean ; Aux : String ; IsAux : Boolean; bCFONBIntern : Boolean = False) : boolean ;
{$IFNDEF EAGLSERVER}
procedure ZeroBlanc ( P : TPanel ) ;
{$ENDIF}
function  MvtToIdent( Q : TQuery ; fb : TFichierBase ; Totale : boolean ) : RMVT ;
function  SQLForIdent( fb : TFichierBase ) : string ;
function  BudgetToIdent ( Q : TQuery ) : RMVT ;
Function  StrToTypCtr(St : String3) : ttCtrPtr ;
{$IFNDEF EAGLSERVER}
procedure ProrateQteAnal ( GS : THGrid ; Lig : integer ; OQ,NQ : Double ; Ind : integer ) ;
{$ENDIF}
Function  ConvertJoker(nom : string): String  ;
function  IncrementeNumeroBudget : longint ;
procedure ECCANA ( T : TDataSet ; CodeD : String3 ) ;
procedure ECCECR ( T : TDataSet ; CodeD : String3 ) ;
procedure RemplirOFromM ( M : RMVT ; ListeSel : TList ) ;
function  EstFormule ( St : String ) : boolean ;
Function  Code2Exige ( Code : String3 ) : TExigeTva ;
Function  FlagEncais ( Exige : TExigeTva ; SurEnc : boolean ) : String ;
{$IFNDEF EAGLSERVER}
Function  RecupDebit ( GS : THGrid ; Lig : integer ; tm : TSorteMontant ) : Double ;
Function  RecupCredit ( GS : THGrid ; Lig : integer ; tm : TSorteMontant ) : Double ;
procedure RecupTotalPivot ( GS : THGrid ; Lig : integer ; var valDeb, valCre : double ) ; // SBO 09/08/2007 FQ 20910 maj solde des comptes erroné en devise
{$ENDIF}
function  CoherQualif ( L : TList ; QQ : String3 ; Ind : integer ) : boolean ;
Procedure InitTableTemporaire(TypeTraitement : String) ;
Function AlimTableTemporaire(O : TOBM ; LeTypeTraitement : String ; smp : tSuiviMP ; ZL1,ZL2 : String) : Boolean ;
{$ENDIF}
{$ENDIF}

Procedure ElementsTvaEnc ( M : RMVT ; ForceFlag : boolean ; ExiConnue : String = ''; AccepteOD : boolean = false ) ;
// ME 04/06/2007 unification de la fonction  ElementsTvaEnc avec  TraitementSurEcr fiche 10537

{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
Procedure ElementsTvaEncRevise ( M : RMVT ; ForceFlag : boolean ; Var InfoImp : TInfoImport ; QFiche : TQFiche; ExiConnue : String = ''; AccepteOD : boolean = false ) ;
{$ENDIF}
{$ENDIF}

{$IFNDEF NOVH}
{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
procedure RenseigneHTByTVA ( CGen : TGGeneral ; RegTVA,CodeTva,CodeTpf : String3 ; SoumisTpf,Achat,EncON : boolean ; Var CoherTva : boolean ) ;
{$ENDIF}
{$ENDIF}
Function  EstJalFact ( Jal : String3 ) : boolean ;
{$ENDIF}

Function  Tva2NumBase ( Code : String3 ) : integer ;
Function  EstCollFact ( Coll : String ) : boolean ;
function  CGetTypeExo ( vDtDateComptable : TDateTime ) : TTypeExo;
procedure AttribParamsNew ( Var FRM : TFRM ; D,C : Double ; T : TTypeExo ) ;
{$IFNDEF EAGLSERVER}
procedure GereNewLigne ( GS : THGrid ) ;
procedure ColorOpposeEuro ( G : THGrid ; ModeOppose,EnDevise : boolean ) ;
Procedure DelTabsSerie ( TA : TTabControl ) ;
{$ENDIF}

procedure Ajoute ( TS : TList ; St : String3 ; Cpte : String17 ; D,C : double ; NumP,NumL : Integer ; DateP : tDateTime ; CptCegid : string = '' ; YtcTiers : string = '' ) ;
procedure AjouteAno ( TS : TList ; vTOBEcr : TOB ; NatGene : string ; Inverse : boolean ) ;
{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
Function AlimLTabCptLu(What : Integer ; Q : TQuery ; L : HTStringList ; ListeCptFaux : TList ; Var CptLu : TCptLu ) : Boolean ;
Procedure ClassToRec(Var C1 : TCptLu ; C2 : TCptLuP) ;
Function MessErreurCompte(What : Integer ; CptLu : TCptLu) : String ;
Function ChercheCptLu(L : HTStringList ; Var CptLu : TCptLu) : Boolean ;
{$ENDIF}
{$ENDIF}

procedure VideTFRM ( var Value : TFRM ) ;

implementation

Uses
{$IFNDEF PGIIMMO}
 ULibecriture ,
{$ENDIF}
 ULibExercice ,
{$IFNDEF CCADM}
{$IFDEF COMPTA}
  EcheMPA,
{$ENDIF}
{$ENDIF}
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
     FichComm,
  {$ENDIF !ERADIO}
{$ENDIF}
   {$IFDEF MODENT1}
   CPProcMetier,
   CPProcGen, 
   {$ENDIF MODENT1}
     UtilSais ;



{$IFNDEF EAGLSERVER}
Function ValD ( GS : THGrid ; Lig : integer ) : double ;
BEGIN
Case GS.TypeSais of
   tsGene,tsTreso : Result:=Valeur(GS.Cells[SA_Debit,Lig])  ;
            tsODA : Result:=Valeur(GS.Cells[OA_Debit,Lig])  ;
            else Result:=Valeur(GS.Cells[SA_Debit,Lig])  ;
   END ;
END ;

Function ValC ( GS : THGrid ; Lig : integer ) : double ;
BEGIN
Case GS.TypeSais of
   tsGene,tsTreso : Result:=Valeur(GS.Cells[SA_Credit,Lig])  ;
            tsODA : Result:=Valeur(GS.Cells[OA_Credit,Lig])  ;
            else Result:=0 ;
   END ;
END ;

Function GetO ( GS : THGrid ; Lig : integer = -1 ) : TOBM ;
BEGIN
if Lig=-1 then Lig:=GS.Row ;
Case GS.TypeSais of
   tsAnal      : Result:=TOBM(GS.Objects[AN_NumL,Lig]) ;
   tsLettrage  : Result:=TOBM(GS.Objects[GS.ColCount-1,Lig]) ;
   tsPointage  : Result:=TOBM(GS.Objects[0,Lig]) ;
   else Result:=TOBM(GS.Objects[SA_Exo,Lig]) ;
   END ;
END ;

procedure SetO ( GS : ThGrid ; O : TOB ; Lig : integer = -1 ) ;
begin
if Lig=-1 then Lig:=GS.Row ;
 Case GS.TypeSais of
   tsAnal      : GS.Objects[AN_NumL,Lig]:=O ;
   tsLettrage  : GS.Objects[GS.ColCount-1,Lig]:=O ;
   tsPointage  : GS.Objects[0,Lig]:=O ;
   else GS.Objects[SA_Exo,Lig]:=O ;
   END ;
end;
{$ENDIF}

{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
{$IFNDEF EAGLSERVER}
function MontantLigne ( GS : THGrid ; Lig : integer ; tm : TSorteMontant ) : double ;
BEGIN
Result:=RecupDebit(GS,Lig,tm) ; if Result=0 then Result:=RecupCredit(GS,Lig,tm) ;
END ;


Function GetChamp( GS : THGrid ; Lig : integer ; Quoi : String ) : Variant ;
BEGIN Result:=GetO(GS,Lig).GetMvt(Quoi) ; END ;

Function GetDouble( GS : THGrid ; Lig : integer ; Quoi : String ) : Double ;
BEGIN Result:=Double(GetChamp(GS,Lig,Quoi)) ; END ;
{$ENDIF}
{$ENDIF}
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 13/11/2001
Modifié le ... : 30/03/2005
Description .. : Retourne le type d'exo en fonction de la date comptable
Suite ........ : SBO 30/03/2005 : Ajout d'un paramètre vBoAuto = True.
Suite ........ : Permet de prendre en compte la date passée en paramètre
Suite ........ : si fonction appelé avec False.
Mots clefs ... : EXERCICE
*****************************************************************}

function CGetTypeExo ( vDtDateComptable : TDateTime ) : TTypeExo;
var
 lStCodeExoCourant : string;
begin

 lStCodeExoCourant := QuelExoDt ( vDtDateComptable ) ;

(* ajout me Fiche 14973 Pb on force l'exercice de référence
  if ( ctxPCL in V_PGI.PGIContexte )                and
    ( ( GetCPExoRef.Code = GetEncours.Code )     or
    ( GetCPExoRef.Code = GetSuivant.Code ) )
 then lStCodeExoCourant   := GetCPExoRef.Code ;
*)

 if lStCodeExoCourant = GetEncours.Code      then result := teEncours
 else if lStCodeExoCourant = GetSuivant.Code then result := teSuivant
 else result := tePrecedent;

end;

procedure AttribParamsNew ( Var FRM : TFRM ; D,C : Double ; T : TTypeExo ) ;
BEGIN
FRM.Deb:=D ; FRM.Cre:=C ;
Case T of
   teEncours : BEGIN
               FRM.DE:=D ; FRM.CE:=C ;
               FRM.DS:=0 ; FRM.CS:=0 ;
               FRM.DP:=0 ; FRM.CP:=0 ;
               END ;
   teSuivant : BEGIN
               FRM.DE:=0 ; FRM.CE:=0 ;
               FRM.DS:=D ; FRM.CS:=C ;
               FRM.DP:=0 ; FRM.CP:=0 ;
               END ;
 tePrecedent : BEGIN
               FRM.DE:=0 ; FRM.CE:=0 ;
               FRM.DS:=0 ; FRM.CS:=0 ;
               FRM.DP:=D ; FRM.CP:=C ;
               END ;
   END ;
END ;


Function EcheMvt ( T : TDataSet ) : Boolean ;
BEGIN Result:=(T.FindField('E_NUMECHE').AsInteger>0) ; END ;

Function AnalMvt ( T : TDataSet ) : Boolean ;
BEGIN Result:=(T.FindField('E_ANA').AsString='X') ; END ;

{$IFNDEF EAGLSERVER}
procedure NewLigne ( GS : THGrid )  ;
BEGIN GS.RowCount:=GS.RowCount+1 ; END ;

procedure GereNewLigne ( GS : THGrid ) ;
BEGIN
if True then
   BEGIN
   if EstRempli(GS,GS.RowCount-1) then NewLigne(GS) else
      if Not EstRempli(GS,GS.RowCount-2) then GS.RowCount:=GS.RowCount-1 ;
   END ;
END ;

Function EstRempli( GS : THGrid ; Lig : integer) : Boolean ;
BEGIN
Case GS.TypeSais of
   tsGene,tsTreso : Result:=(GS.Cells[SA_Gen,Lig]<>'') or (GS.Cells[SA_Aux,Lig]<>'') ;
           tsAnal : Result:=GS.Cells[AN_Sect,Lig]<>''  ;
            tsODA : Result:=GS.Cells[OA_Sect,Lig]<>''  ;
   else Result:=False ;
   END ;
END ;
{$ENDIF}

{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
{$IFNDEF EAGLSERVER}
Function isBQE ( GS : THGrid ; Lig : integer ) : boolean ;
Var Nat : String3 ;
BEGIN
Result:=True ;
if ((GS.Cells[SA_Gen,Lig]<>'') and (GS.Objects[SA_Gen,Lig]<>Nil)) then
    BEGIN
    Nat:=GetGGeneral(GS,Lig).NatureGene ;
    if ((Nat='BQE') or (Nat='CAI')) then Exit ;
    END ;
Result:=False ;
END ;

Function isTiers ( GS : THGrid ; Lig : integer ) : boolean ;
Var Nat : String3 ;
BEGIN
Result:=True ;
if ((GS.Cells[SA_Aux,Lig]<>'') and (GS.Objects[SA_Aux,Lig]<>Nil)) then Exit ;
if ((GS.Cells[SA_Gen,Lig]<>'') and (GS.Objects[SA_Gen,Lig]<>Nil)) then
    BEGIN
    Nat:=GetGGeneral(GS,Lig).NatureGene ;
    if ((Nat='TID') or (Nat='TIC')) then Exit ;
    END ;
Result:=False ;
END ;

Function isHT ( GS : THGrid ; Lig : integer ; OkTVA : boolean ) : boolean ;
Var Nat  : String3 ;
    CGen : TGGeneral ;
BEGIN
isHT:=False ;
if Not EstRempli(GS,Lig) then Exit ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
Nat:=CGen.NatureGene ;
if ((Nat='CHA') or (Nat='PRO')) then isHT:=True else
   if ((Nat='IMO') and (OkTVA)) then isHT:=True ;
END ;

function ArretZone(Col : integer ; Var St : String ; MAJ : boolean ) : boolean ;
Var Ind : integer ;
BEGIN
Ind:=-1 ; ArretZone:=False ;
if Col=SA_Gen then Ind:=1 else if Col=SA_Aux then Ind:=2 else if Col=SA_RefI then Ind:=3 else
   if Col=SA_Lib then Ind:=4 else if Col=SA_Debit then Ind:=5 else if Col=SA_Credit then Ind:=6 ;
if Ind=-1 then Exit ;
ArretZone:=St[Ind]='X' ; if MAJ then St[Ind]:='-' ;
END ;

Function TrouveLigCompte ( GS : THGrid ; Col,FromLig : integer ; Cpte : String17 ) : integer ;
Var i : integer ;
BEGIN
Result:=-1 ;
for i:=GS.RowCount-2 downto FromLig+1 do
    BEGIN
    if ((GS.Cells[Col,i]=Cpte) and (GS.Objects[Col,i]<>Nil)) then BEGIN Result:=i ; Exit ; END ;
    END ;
END ;

Function TrouveLigHT ( GS : THGrid ; FromLig : integer ; OkTVA : boolean ) : integer ;
Var i : integer ;
BEGIN
Result:=-1 ;
for i:=FromLig+1 to GS.RowCount-2 do if IsHT(GS,i,OkTva) then BEGIN Result:=i ; Exit ; END ;
END ;

Function TrouveLigTiers ( GS : THGrid ; FromLig : integer) : integer ;
Var i : integer ;
BEGIN
Result:=-1 ;
for i:=FromLig+1 to GS.RowCount-2 do if isTiers(GS,i) then BEGIN Result:=i ; Exit ; END ;
END ;

Function TrouveLigTVA ( GS : THGrid ; FromLig : integer ) : integer ;
Var i,j      : integer ;
    CP,CA,CF : String ;
    CGen     : TGGeneral ;
BEGIN
TrouveLigTVA:=0 ;
for i:=FromLig+1 to GS.RowCount-2 do if ((GS.Cells[SA_Gen,i]<>'') and (GS.Objects[SA_Gen,i]<>Nil)) then
    BEGIN
    CP:=GS.Cells[SA_Gen,i] ;
    for j:=1 to GS.RowCount-2 do
        BEGIN
        CGen:=GetGGeneral(GS,j) ; if CGen=Nil then Continue ;
        CA:=CGen.CpteTVA ; CF:=CGen.CpteTPF ;
        if ((CP=CA) or (CP=CF)) then BEGIN TrouveLigTVA:=i ; Exit ; END ;
        END ;
    END ;
END ;


procedure SwapOutil(Outils : TPanel) ;
BEGIN
if Outils.Align=alTop then Outils.Align:=alNone else Outils.Align:=alTop ;
END ;

procedure NumeroteLignes ( GS : THGrid ) ;
Var i : integer ;
    O : TOBM ;
BEGIN
for i:=1 to GS.RowCount-1 do
    BEGIN
    Case GS.TypeSais of
     tsGene,tsTreso : BEGIN
                      GS.Cells[SA_NumL,i]:=IntToStr(i) ;
                      O:=GetO(GS,i) ; if O<>Nil then O.PutMvt('E_NUMLIGNE',i) ;
                      END ;
             tsAnal : BEGIN
                      GS.Cells[AN_NumL,i]:=IntToStr(i) ; O:=GetO(GS,i) ;
                      if O<>Nil then O.PutMvt('Y_NUMVENTIL',i) ;
                      END ;
              tsODA : BEGIN
                      GS.Cells[OA_NumL,i]:=IntToStr(i) ; O:=GetO(GS,i) ;
                      if O<>Nil then O.PutMvt('Y_NUMVENTIL',i) ;
                      END ;
     End ;
    END ;
END ;

procedure AlloueMvt ( GS : THGrid ; Ecr : TTypeEcr ; Lig : integer ; Init : boolean ) ;
Var OBM : TOBM ;
BEGIN
if True then
   BEGIN
   if GS.Objects[SA_Exo,Lig]<>Nil then exit ;
   if GS.Cells[SA_Exo,Lig]='' then Exit ;
   OBM:=TOBM.Create(Ecr,'',Init) ; GS.Objects[SA_Exo,Lig]:=TObject(OBM) ;
   END ;
END ;

Function RecupDebit ( GS : THGrid ; Lig : integer ; tm : TSorteMontant ) : Double ;
Var OBM : TOBM ;
BEGIN
Result:=0 ; OBM:=GetO(GS,Lig) ; if OBM=Nil then Exit ;
Case tm of
   tsmPivot  : Result:=OBM.GetMvt('E_DEBIT') ;
   tsmDevise : Result:=OBM.GetMvt('E_DEBITDEV') ;
   END ;
END ;

Function RecupCredit ( GS : THGrid ; Lig : integer ; tm : TSorteMontant ) : Double ;
Var OBM : TOBM ;
BEGIN
Result:=0 ; OBM:=GetO(GS,Lig) ; if OBM=Nil then Exit ;
Case tm of
   tsmPivot  : Result:=OBM.GetMvt('E_CREDIT') ;
   tsmDevise : Result:=OBM.GetMvt('E_CREDITDEV') ;
   END ;
END ;

// SBO 09/08/2007 FQ 20910 maj solde des comptes erroné en devise
procedure RecupTotalPivot ( GS : THGrid ; Lig : integer ; var valDeb, valCre : double ) ;
var lModR : TMOD ;
begin

  valDeb := 0 ;
  valCre := 0 ;

  // Pour le Multi-échéance, il faut récupérer le montant total de la ligne
  lModR   := TMOD(GS.Objects[SA_NumP,Lig]) ;
  if Assigned(lModR) and (lModR.ModR.NbEche>1) then
    begin
    if ValD(GS,Lig) <> 0
      then valDeb := lModR.ModR.TotalAPayerP
      else valCre := lModR.ModR.TotalAPayerP ;
    end
  else
    begin
    valDeb := RecupDebit(GS,Lig,tsmPivot) ;
    valCre := RecupCredit(GS,Lig,tsmPivot) ;
    end ;

end;
// Fin FQ 20910

procedure Desalloue( GS : THGrid ; LaCol,LaLig : Longint ) ;
BEGIN
if GS.Objects[LaCol,LaLig]=Nil Then Exit ;
if True then
   BEGIN
   if LaCol=SA_Gen then BEGIN TGGeneral(GS.Objects[LaCol,LaLig]).Free ; GS.Objects[LaCol,LaLig]:=Nil ; END else
   if LaCol=SA_Aux then BEGIN TGTiers(GS.Objects[LaCol,LaLig]).Free ; GS.Objects[LaCol,LaLig]:=Nil ; END ;
   END ;
END ;

Function LaMeme ( GS : THGrid ; L : integer) : Boolean ;
Var Col : integer ;
BEGIN
Result:=False ;
if GS.TypeSais in [tsAnal] then Col:=AN_Sect else if GS.TypeSais=tsODA then Col:=OA_Sect else Exit ;
if GS.Objects[Col,L]=Nil then Exit ; if GS.Cells[Col,L]='' then Exit ;
if GetGSect(GS,L).SECT<>GS.Cells[Col,L] then Exit ;
Result:=True ;
END ;

Function LeMeme ( GS : THGrid ; C,L : integer) : Boolean ;
BEGIN
LeMeme:=False ;
if GS.Objects[C,L]=Nil then Exit ; if GS.Cells[C,L]='' then Exit ;
if True then
   BEGIN
   if C=SA_Gen then BEGIN if uppercase(GetGGeneral(GS,L).GENERAL)<>uppercase(GS.Cells[C,L]) then exit ; END else
      if C=SA_Aux then BEGIN if uppercase(GetGTiers(GS,L).AUXI)<>uppercase(GS.Cells[C,L]) then exit ; END ;
   END ;
LeMeme:=True ;
END ;

Function TrouveSens ( GS : THGrid ; Nat : String3 ; Lig : integer ; CommeBQE : boolean = False ) : byte ;
Var CGen : TGGeneral ;
BEGIN
TrouveSens:=3 ;
if GS.Cells[SA_Gen,Lig]='' then exit ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then Exit ;
TrouveSens:=QuelSens(Nat,CGen.NatureGene,CGen.Sens,CommeBQE) ;
END ;

Procedure AttribGen ( GS : THGrid ; LigC,Lig : integer ) ;
Var CGen,GG : TGGeneral ;
BEGIN
GG:=GetGGeneral(GS,LigC) ; CGen:=TGGeneral.Create('') ;
CGen.General:=GG.General ;
CGen.Libelle:=GG.Libelle       ; CGen.NatureGene:=GG.NatureGene ; CGen.RegimeTVA:=GG.RegimeTVA   ; CGen.Collectif:=GG.Collectif ;
CGen.Lettrable:=GG.Lettrable   ; CGen.Ventilable:=GG.Ventilable ; CGen.TotalDebit:=GG.TotalDebit ; CGen.TotalCredit:=GG.TotalCredit ;
CGen.Budgene:=GG.Budgene       ; CGen.Tva:=GG.Tva               ; CGen.Tpf:=GG.Tpf ;
CGen.SoumisTPF:=GG.SoumisTPF   ; CGen.ModeRegle:=GG.ModeRegle   ; CGen.CpteTVA:=GG.CpteTVA       ; CGen.CpteTPF:=GG.CpteTPF ;
CGen.TauxTVA:=GG.TauxTVA       ; CGen.TauxTPF:=GG.TauxTPF       ; CGen.Pointable:=GG.Pointable   ;
CGen.Ferme:=GG.Ferme           ; CGen.Sens:=GG.Sens             ; CGen.Tva_Encaissement:=GG.Tva_Encaissement ;
CGen.QQ1:=GG.QQ1               ; CGen.QQ2:=GG.QQ2               ; CGen.TvaSurEncaissement:=GG.TvaSurencaissement ;
CGen.Abrege:=GG.Abrege         ; CGen.SuiviTreso:=GG.SuiviTreso ; CGen.Confidentiel:=GG.Confidentiel ;
CGen.CodeConso:=GG.CodeConso   ;
GS.Objects[SA_Gen,Lig]:=TObject(CGen) ;
END ;

Procedure AttribAux ( GS : THGrid ; LigC,Lig : integer ) ;
Var CAux,TT : TGTiers ;
BEGIN
TT:=GetGTiers(GS,LigC) ; CAux:=TGTiers.Create('') ;
CAux.Auxi:=TT.Auxi                   ; CAux.Libelle:=TT.Libelle ;
CAux.JourPaiement1:=TT.JourPaiement1 ; CAux.JourPaiement2:=TT.JourPaiement2 ;
CAux.Collectif:=TT.Collectif         ; CAux.NatureAux:=TT.NatureAux ;
CAux.TotalDebit:=TT.TotalDebit       ; CAux.TotalCredit:=TT.TotalCredit ;
CAux.Lettrable:=TT.Lettrable         ; CAux.ModeRegle:=TT.ModeRegle ;
CAux.RegimeTva:=TT.RegimeTVA         ; CAux.SoumisTPF:=TT.SoumisTPF ;
CAux.Abrege:=TT.Abrege               ; CAux.Tva_Encaissement:=TT.Tva_Encaissement ;
CAux.Confidentiel:=TT.Confidentiel   ; CAux.Ferme:=TT.Ferme ;
CAux.Devise:=TT.Devise               ; CAux.MultiDevise:=TT.MultiDevise ;
CAux.CodeConso:=TT.CodeConso         ;
GS.Objects[SA_Aux,Lig]:=TObject(CAux) ;
END ;
{$ENDIF}


{$IFNDEF EAGLSERVER}
procedure PositionneTz( Cache : THCpteEdit ; ip : integer ) ;
BEGIN
Case ip of
  1 : Cache.ZoomTable:=tzSection  ; 2 : Cache.ZoomTable:=tzSection2 ;
  3 : Cache.ZoomTable:=tzSection3 ; 4 : Cache.ZoomTable:=tzSection4 ;
  5 : Cache.ZoomTable:=tzSection5 ;
  END ;
END ;
{$ENDIF}
{$ENDIF}
{$ENDIF}




function FabricReqComp ( fb : TFichierBase  ; Var FRM : TFRM ) : String ;
Var SQL   : String ;
    i     : integer ;
BEGIN
Result:='' ; SQL:='' ;
Case fb of
   fbGene : BEGIN
            SQL:=' AND G_LETTRABLE="'+FRM.G_L+'" AND G_COLLECTIF="'+FRM.G_C+'" AND G_NATUREGENE="'+FRM.G_N+'"' ;
            for i:=1 to MaxAxe do SQL:=SQL+' AND G_VENTILABLE'+Inttostr(i)+'="'+FRM.G_V[i]+'" ' ;
            END ;
   fbAux  : SQL:=' AND T_LETTRABLE="'+FRM.T_L+'" AND T_NATUREAUXI="'+FRM.T_N+'"' ;
   End ;
Result:=SQL ;
END ;



function ExecReqMAJBilanAno( FRM : TFRM ) : boolean ;
var
 lQ               : TQuery ;
 lRdDebit         : double ;
 lRdCredit        : double ;
 lRdDebitDev      : double ;
 lRdCreditDev     : double ;
 lInNumLigne      : integer ;
 lInNumeroPiece   : integer ;
 lDtDateModif     : TDateTime ;
 lDtDATECOMPTABLE : TDateTime ;
 lStJal           : string ;
 lTOB             : TOB ;
 lRdSolde         : double ;
 lRdSoldeDev      : double ;
 lRdTaux          : double ;
 lRdCotation      : double ;
begin

 result := true ;

 if ( Arrondi(FRM.Deb,4) = 0 ) and ( Arrondi(FRM.Cre,4) = 0 ) then exit ;

 lRdCredit    := 0 ;
 lRdDebit     := 0 ;
 lRdCreditDev := 0 ;
 lRdDebitDev  := 0 ;


 lQ := OpenSQL ('select e_debit,e_credit,e_debitdev,e_creditdev,e_datemodif,e_numeropiece,e_numligne from ecriture ' +
                'where e_exercice="'      + ctxExercice.Suivant.Code   + '" and ' +
                ' e_journal="'            + GetParamSoc('SO_JALOUVRE') + '" and ' +
                ' e_etablissement="'      + FRM.Etabliss               + '" and ' +
                ' e_devise="'             + FRM.Dev                    + '" and ' +
                ' e_general="'            + FRM.Cpt                    + '" and ' +
                ' e_auxiliaire="'         + FRM.Axe                    + '" ' , true,-1,'',true) ;

  if not lQ.EOF then
   begin

     lRdSolde    := Arrondi( ( lQ.FindField('E_CREDIT').asFloat + FRM.Cre ) - ( lQ.FindField('E_DEBIT').asFloat + FRM.Deb ) , V_PGI.OkDecV ) ;
     lRdSoldeDev := Arrondi( ( lQ.FindField('E_CREDITDEV').asFloat + FRM.CreDev ) - ( lQ.FindField('E_DEBITDEV').asFloat + FRM.DebDev ) , V_PGI.OkDecV ) ;

     if lRdSolde < 0 then
      begin
       lRdCredit    := lRdSolde * (-1) ;
       lRdCreditDev := lRdSoldeDev * (-1) ;
      end
       else
        begin
         lRdDebit    := lRdSolde ;
         lRdDebitDev := lRdSoldeDev ;
        end ;

     if lRdSoldeDev <> 0 then
      lRdTaux := Arrondi( ( lRdSolde  / lRdSoldeDev ) , V_PGI.OkDecV )
       else
        lRdTaux := 1 ;
        
     if V_PGI.TauxEuro <> 0 then
       lRdCotation := lRdTaux / V_PGI.TauxEuro
        else
         lRdCotation := 1 ;

     lDtDateModif    := lQ.FindField('E_DATEMODIF').asDateTime ;
     lInNumeroPiece  := lQ.FindField('E_NUMEROPIECE').asInteger ;
     lInNumLigne     := lQ.FindField('E_NUMLIGNE').asInteger ;
     Ferme(lQ) ;

     result := ExecuteSQL('UPDATE ECRITURE Set e_debit='      + StrFPoint(lRdCredit)       +
                             ', e_credit='                    + StrFPoint(lRdDebit)        +
                             ', e_debitdev='                  + StrFPoint(lRdCreditDev)    +
                             ', e_creditdev='                 + StrFPoint(lRdDebitDev)     +
                             ', e_tauxdev='                   + StrFPoint(lRdTaux)         +
                             ', e_cotation='                  + StrFPoint(lRdCotation)     +
                             ' where e_exercice="'            + ctxExercice.Suivant.Code   + '" and ' +
                             ' e_journal="'                   + GetParamSoc('SO_JALOUVRE') + '" and ' +
                             ' e_numeropiece='                + intToStr(lInNumeroPiece)   + ' and '  +
                             ' e_datemodif="'                 + UsTime(lDtDateModif)       + '" and ' +
                             ' e_numligne='                   + intToStr(lInNumLigne)   ) = 1 ;

   end
    else
     begin

      Ferme(lQ) ;

      lStJal           := GetParamSocSecur('SO_JALOUVRE' , '' ) ;
      lDtDATECOMPTABLE := ctxExercice.Suivant.Deb ;
      lRdSoldeDev      := FRM.DebDev + FRM.CreDev ;

      if ( lRdSoldeDev ) <> 0 then
       lRdTaux := Arrondi( ( ( FRM.Deb + FRM.Cre ) / ( lRdSoldeDev ) ) , V_PGI.OkDecV )
        else
         lRdTaux := 1 ;

      if V_PGI.TauxEuro <> 0 then
       lRdCotation := lRdTaux / V_PGI.TauxEuro
        else
         lRdCotation := 1 ;

      lQ := OpenSQL ('select max(e_numligne) N,max(e_numeropiece) P from ecriture ' +
                     'where e_exercice="'      + ctxExercice.Suivant.Code   + '" and ' +
                     ' e_journal="'            + lStJal                     + '" and ' +
                     ' e_etablissement="'      + FRM.Etabliss               + '" and ' +
                     ' e_devise="'             + FRM.Dev                    + '" '
                      , true,-1,'',true) ;
      if lQ.FindField('P').asInteger > 0 then
       begin
        lInNumeroPiece  := lQ.FindField('P').asInteger ;
        lInNumLigne     := lQ.FindField('N').asInteger + 1 ;
        Ferme(lQ) ;
       end
        else
         begin
          Ferme(lQ) ;
          lInNumeroPiece   := GetNewNumJal( lStJal , true , lDtDATECOMPTABLE ) ;
          lInNumLigne      := 1 ;
         end ;

      lTOB             := TOB.Create('ECRITURE',nil,-1) ;
{$IFNDEF PGIIMMO}
      CPutDefautEcr(lTOB) ;
{$ENDIF}
      lTOB.PutValue('E_EXERCICE'                          , ctxExercice.Suivant.Code ) ;
      lTOB.PutValue('E_JOURNAL'                           , lStJal ) ;
      lTOB.PutValue('E_DATECOMPTABLE'                     , lDtDATECOMPTABLE ) ;
      lTOB.PutValue('E_NUMEROPIECE'                       , lInNumeroPiece ) ;
      lTOB.PutValue('E_NUMLIGNE'                          , lInNumLigne ) ;
      lTOB.PutValue('E_GENERAL'                           , FRM.Cpt ) ;
      lTOB.PutValue('E_AUXILIAIRE'                        , FRM.Axe ) ;
      lTOB.PutValue('E_DEBIT'                             , Arrondi(FRM.Deb,V_PGI.OkDecV) ) ;
      lTOB.PutValue('E_CREDIT'                            , Arrondi(FRM.Cre,V_PGI.OkDecV) ) ;
      lTOB.PutValue('E_DEBITDEV'                          , Arrondi(FRM.DebDev,V_PGI.OkDecV) ) ;
      lTOB.PutValue('E_CREDITDEV'                         , Arrondi(FRM.CreDev,V_PGI.OkDecV) ) ;
      lTOB.PutValue('E_NATUREPIECE'                       ,'OD' ) ;
      lTOB.PutValue('E_ECRANOUVEAU'                       ,'OAN' ) ;
      lTOB.PutValue('E_MODESAISIE'                        , '-' ) ;
      lTOB.PutValue('E_TRESOSYNCHRO'                      , 'RIE' ) ;
      lTOB.PutValue('E_QUALIFPIECE'                       , 'N' ) ;
      lTOB.PutValue('E_VALIDE'                            , 'X' ) ;
      lTOB.PutValue('E_DATECREATION'                      , lDtDATECOMPTABLE ) ;
      lTOB.PutValue('E_DATEMODIF'                         , Now ) ;
      lTOB.PutValue('E_SOCIETE'                           , GetParamSocSecur('SO_SOCIETE','') ) ;
      lTOB.PutValue('E_ETABLISSEMENT'                     , FRM.Etabliss ) ;
      lTOB.PutValue('E_DEVISE'                            , FRM.Dev ) ;
      lTOB.PutValue('E_TAUXDEV'                           , lRdTaux ) ;
      lTOB.PutValue('E_COTATION'                          , lRdCotation ) ;
      lTOB.PutValue('E_IO'                                , '-' ) ;
{$IFNDEF PGIIMMO}
      CSetPeriode(lTOB) ;
      CSetCotation(lTOB) ;
{$ENDIF}
      lTOB.InsertDB(nil) ;
      lTOB.Free ;

     result := true ;

  end ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 09/06/2006
Modifié le ... :   /  /    
Description .. : - LG  - 09/06/2006 - FB 18349 - coorection de sl maj des
Suite ........ : ano pour les compte de charges et produits
Mots clefs ... :
*****************************************************************}
function ExecReqMAJAnoAno( var FRM : TFRM ) : boolean ;
var
 lQ               : TQuery ;
 lRdDebit         : double ;
 lRdCredit        : double ;
 lRdSolde         : double ;
 lRdDebitDev      : double ;
 lRdCreditDev     : double ;
 lRdSoldeDev      : double ;
 lInNumLigne      : integer ;
 lInNumeroPiece   : integer ;
 //lStCompteSolde   : string ;
 lDtDateModif     : TDateTime ;
 lDtDATECOMPTABLE : TDateTime ;
 lTOB             : TOB ;
 lStJal           : string ;
 lRdTaux          : double ;
 lRdCotation      : double ;
begin

 result := true ;

 if ( Arrondi(FRM.Deb,4) = 0 ) and ( Arrondi(FRM.Cre,4) = 0 ) then exit ;

 result := false ;

 lRdCredit    := 0 ;
 lRdDebit     := 0 ;
 lRdCreditDev := 0 ;
 lRdDebitDev  := 0 ;

 lQ := OpenSQL ('select e_debit,e_credit,e_debitdev,e_creditdev,e_datemodif,e_numeropiece,e_numligne from ecriture ' +
                'where e_exercice="'      + ctxExercice.Suivant.Code   + '" and ' +
                ' e_journal="'            + GetParamSoc('SO_JALOUVRE') + '" and ' +
                ' e_etablissement="'      + FRM.Etabliss               + '" and ' +
                ' e_devise="'             + FRM.Dev                    + '" and ' +
                ' ( ( e_general="'        + GetParamSoc('SO_FERMEBEN') + '" ) or '  +
                ' ( e_general="'          + GetParamSoc('SO_FERMEPERTE') + '" ) ) '
                , true,-1,'',true) ;

  if not lQ.EOF then
   begin

     lRdSolde    := Arrondi( ( lQ.FindField('E_CREDIT').asFloat + FRM.Cre ) - ( lQ.FindField('E_DEBIT').asFloat + FRM.Deb ) , V_PGI.OkDecV ) ;
     lRdSoldeDev := Arrondi( ( lQ.FindField('E_CREDITDEV').asFloat + FRM.CreDev ) - ( lQ.FindField('E_DEBITDEV').asFloat + FRM.DebDev ) , V_PGI.OkDecV ) ;

     if lRdSolde > 0 then
      begin
       lRdCredit      := lRdSolde ;
       lRdCreditDev   := lRdSoldeDev ;
       FRM.Cpt        := GetParamSocSecur('SO_FERMEBEN', '') ;
      end
       else
        begin
         lRdDebit       := lRdSolde * (-1) ;
         lRdDebitDev    := lRdSoldeDev * (-1) ;
         FRM.Cpt        := GetParamSocSecur('SO_FERMEPERTE', '') ;
        end ;

      if lRdSoldeDev <> 0 then
       lRdTaux := Arrondi( ( lRdSolde / lRdSoldeDev ) , V_PGI.OkDecV )
        else
         lRdTaux := 1 ;

      if V_PGI.TauxEuro <> 0 then
       lRdCotation := lRdTaux / V_PGI.TauxEuro
        else
         lRdCotation := 1 ;

      lDtDateModif    := lQ.FindField('E_DATEMODIF').asDateTime ;
      lInNumeroPiece  := lQ.FindField('E_NUMEROPIECE').asInteger ;
      lInNumLigne     := lQ.FindField('E_NUMLIGNE').asInteger ;
      Ferme(lQ) ;

      result := ExecuteSQL('UPDATE ECRITURE Set e_debit='      + StrFPoint(lRdDebit)        +
                              ', e_credit='                    + StrFPoint(lRdCredit)       +
                              ',e_debitdev='                   + StrFPoint(lRdDebitDev)     +
                              ', e_creditdev='                 + StrFPoint(lRdCreditDev)    +
                              ', e_tauxdev='                   + StrFPoint(lRdTaux)         +
                              ', e_cotation='                  + StrFPoint(lRdCotation)     +
                              ', e_general="'                  + FRM.Cpt                    + '" '     +
                              ' where e_exercice="'            + ctxExercice.Suivant.Code   + '" and ' +
                              ' e_journal="'                   + GetParamSoc('SO_JALOUVRE') + '" and ' +
                              ' e_numeropiece='                + intToStr(lInNumeroPiece)   + ' and '  +
                              ' e_datemodif="'                 + UsTime(lDtDateModif)       + '" and ' +
                              ' e_numligne='                   + intToStr(lInNumLigne)   ) = 1 ;

   end
    else
     begin

      Ferme(lQ) ;

      lDtDATECOMPTABLE := ctxExercice.Suivant.Deb ;
      lStJal           := GetParamSocSecur('SO_JALOUVRE' , '' ) ;

      lQ := OpenSQL ('select max(e_numligne) N,max(e_numeropiece) P from ecriture ' +
                     'where e_exercice="'      + ctxExercice.Suivant.Code   + '" and ' +
                     ' e_journal="'            + GetParamSoc('SO_JALOUVRE') + '" and ' +
                     ' e_etablissement="'      + FRM.Etabliss               + '" and ' +
                     ' e_devise="'             + FRM.Dev                    + '" '
                      , true,-1,'',true) ;

      if lQ.FindField('P').asInteger > 0 then
       begin
        lInNumeroPiece  := lQ.FindField('P').asInteger ;
        lInNumLigne     := lQ.FindField('N').asInteger + 1 ;
        Ferme(lQ) ;
       end
        else
         begin
          Ferme(lQ) ;
          lInNumeroPiece   := GetNewNumJal( lStJal , true , lDtDATECOMPTABLE ) ;
          lInNumLigne      := 1 ;
         end ;

      lDtDATECOMPTABLE := ctxExercice.Suivant.Deb ;


      if  Arrondi( FRM.Cre-FRM.Deb , V_PGI.OkDecV ) > 0 then
       FRM.Cpt := GetParamSocSecur('SO_FERMEBEN', '')
        else
         FRM.Cpt := GetParamSocSecur('SO_FERMEPERTE', '') ;

      lRdSolde := FRM.CreDev + FRM.DebDev ;
      if lRdSolde <> 0 then
       lRdTaux := Arrondi( ( FRM.Deb + FRM.Cre ) / ( lRdSolde ) , V_PGI.OkDecV )
        else
         lRdTaux := 1 ;
         
      if V_PGI.TauxEuro <> 0 then
       lRdCotation := lRdTaux / V_PGI.TauxEuro
        else
         lRdCotation := 1 ;


      lTOB     := TOB.Create('ECRITURE',nil,-1) ;
{$IFNDEF PGIIMMO}
      CPutDefautEcr(lTOB) ;
{$ENDIF}
      lTOB.PutValue('E_EXERCICE'                          , ctxExercice.Suivant.Code ) ;
      lTOB.PutValue('E_JOURNAL'                           , lStJal ) ;
      lTOB.PutValue('E_DATECOMPTABLE'                     , lDtDATECOMPTABLE ) ;
      lTOB.PutValue('E_NUMEROPIECE'                       , lInNumeroPiece ) ;
      lTOB.PutValue('E_NUMLIGNE'                          , lInNumLigne ) ;
      lTOB.PutValue('E_GENERAL'                           , FRM.Cpt ) ;
      lTOB.PutValue('E_AUXILIAIRE'                        , '' ) ;
      lTOB.PutValue('E_DEBIT'                             , Arrondi(FRM.Deb,V_PGI.OkDecV) ) ;
      lTOB.PutValue('E_CREDIT'                            , Arrondi(FRM.Cre,V_PGI.OkDecV) ) ;
      lTOB.PutValue('E_DEBITDEV'                          , Arrondi(FRM.DebDev,V_PGI.OkDecV) ) ;
      lTOB.PutValue('E_CREDITDEV'                         , Arrondi(FRM.CreDev,V_PGI.OkDecV) ) ;
      lTOB.PutValue('E_NATUREPIECE'                       ,'OD' ) ;
      lTOB.PutValue('E_ECRANOUVEAU'                       ,'OAN' ) ;
      lTOB.PutValue('E_MODESAISIE'                        , '-' ) ;
      lTOB.PutValue('E_TRESOSYNCHRO'                      , 'RIE' ) ;
      lTOB.PutValue('E_QUALIFPIECE'                       , 'N' ) ;
      lTOB.PutValue('E_VALIDE'                            , 'X' ) ;
      lTOB.PutValue('E_DATECREATION'                      , lDtDATECOMPTABLE ) ;
      lTOB.PutValue('E_DATEMODIF'                         , Now ) ;
      lTOB.PutValue('E_SOCIETE'                           , GetParamSocSecur('SO_SOCIETE','') ) ;
      lTOB.PutValue('E_ETABLISSEMENT'                     , FRM.Etabliss ) ;
      lTOB.PutValue('E_DEVISE'                            , FRM.Dev ) ;
      lTOB.PutValue('E_TAUXDEV'                           , lRdTaux ) ;
      lTOB.PutValue('E_COTATION'                          , lRdCotation ) ;
      lTOB.PutValue('E_IO'                                , '-' ) ;
      // CEGID V9
      lTOB.PutValue('E_DATEORIGINE', iDate1900) ;
      lTOB.PutValue('E_DATEPER',iDate1900) ;
      lTOB.PutValue('E_ENTITY',0 );
      lTOB.PutValue('E_REFGUID','') ;
      // -----------------
{$IFNDEF PGIIMMO}
      CSetPeriode(lTOB) ;
      CSetCotation(lTOB) ;
{$ENDIF}
      lTOB.InsertDB(nil) ;
      lTOB.Free ;

     result := true ;

  end ;

 ExecReqMAJ ( fbGene, false , false , FRM ) ;

 FRM.Cpt := lStJal ;

 ExecReqMAJ ( fbJal, false , false , FRM ) ;


end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/07/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function ExecReqMAJBDS( vTobListeBDS : Tob; vFRM : TFRM ) : Boolean ;
var lQuery       : TQuery;
    lTob         : Tob;
    i            : integer;
    lStCodeBal   : string;
    lRdDebit     : double;
    lRdCredit    : double;
    lRdSolde     : double;
    lRdDebitDev  : double;
    lRdCreditDev : double;
    lRdSoldeDev  : double;
begin
  Result := True;
  if ( Arrondi(vFRM.Deb, V_PGI.OkDecV) = 0 ) and ( Arrondi(vFRM.Cre, V_PGI.OkDecV) = 0) then Exit;
  Result := False;

  for i := 0 to vTobListeBDS.Detail.Count - 1 do
  begin
    lStCodeBal := vTobListeBDS.Detail[i].GetString('BSI_CODEBAL');
    lRdCredit    := 0 ;
    lRdDebit     := 0 ;
    lRdCreditDev := 0 ;
    lRdDebitDev  := 0 ;

    lQuery := OpenSQL ('SELECT BSE_DEBIT, BSE_CREDIT, BSE_DEBITDEV, BSE_CREDITDEV ' +
              'FROM CBALSITECR WHERE ' +
              'BSE_CODEBAL = "' + lStCodeBal + '" AND ' +
              'BSE_COMPTE1 = "' + vFRM.Cpt + '"', True,-1,'',true);

    if not lQuery.Eof then
    begin
      lRdSolde    := Arrondi((lQuery.FindField('BSE_CREDIT').asFloat + vFRM.Cre ) -
                             (lQuery.FindField('BSE_DEBIT').asFloat + vFRM.Deb ), V_PGI.OkDecV);

      lRdSoldeDev := Arrondi((lQuery.FindField('BSE_CREDITDEV').asFloat + vFRM.CreDev ) -
                             (lQuery.FindField('BSE_DEBITDEV').asFloat + vFRM.DebDev ), V_PGI.OkDecV);

      if lRdSolde > 0 then
      begin
        lRdCredit      := lRdSolde ;
        lRdCreditDev   := lRdSoldeDev ;
      end
      else
      begin
        lRdDebit       := lRdSolde * (-1) ;
        lRdDebitDev    := lRdSoldeDev * (-1) ;
      end ;

      Result := ExecuteSQL('UPDATE CBALSITECR SET ' +
                           'BSE_DEBIT     = ' + StrFPoint(lRdDebit) + ', ' +
                           'BSE_CREDIT    = ' + StrFPoint(lRdCredit) + ', ' +
                           'BSE_DEBITDEV  = ' + StrFPoint(lRdDebitDev) + ', ' +
                           'BSE_CREDITDEV = ' + StrFPoint(lRdCreditDev)+ ' WHERE ' +
                           'BSE_CODEBAL = "' + lStCodeBal + '" AND ' +
                           'BSE_COMPTE1 = "' + vFRM.Cpt + '"') = 1;
    end // if not lQuery.Eof
    else
    begin
      lTob := TOB.Create('CBALSITECR', nil, -1);
      lTob.InitValeurs;
      lTob.SetString('BSE_CODEBAL',   lStCodeBal);
      lTob.Setstring('BSE_COMPTE1',   vFRM.Cpt);
      lTob.Setstring('BSE_COMPTE2',   '');
      lTob.SetDouble('BSE_DEBIT',     Arrondi(vFRM.Deb, V_Pgi.OkDecV));
      lTob.SetDouble('BSE_CREDIT',    Arrondi(vFRM.Cre, V_Pgi.OkDecV));
      lTob.SetDouble('BSE_DEBITDEV',  Arrondi(vFRM.DebDev, V_Pgi.OkDecV));
      lTob.SetDouble('BSE_CREDITDEV', Arrondi(vFRM.CreDev, V_Pgi.OkDecV));
      lTob.InsertDB(nil);
      lTob.Free;
      Result := True;
    end;

    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 09/06/2006
Modifié le ... : 13/12/2006
Description .. : - LG  - 09/06/2006 - FB 18349 - correction de sl maj des
Suite ........ : ano pour les compte de charges et produits
Suite ........ : - GCO - 03/07/2006 - Ajout des BDS Dynamiques
Suite ........ : - LG - 06/12/2006 - on lancait la maj des ano dyna. si on 
Suite ........ : cochait balance dyna uniqument
Suite ........ : - LG - 13/12/2006 - FB 19335 - on test l'exercice avant de 
Suite ........ : mettre a jour les anao dyna pb qd on demandait les 
Suite ........ : balances dyna et les ano et que l'on saisisait sur n+1
Mots clefs ... :
*****************************************************************}
function ExecReqMAJAno( TheList : TList ) : boolean ;
{$IFNDEF PGIIMMO}
var lFRM         : TFRM ;
    i            : integer ;
    lBoBDSDyna   : boolean ;
    lBoAnoDyna   : boolean ;
    lTobListeBDS : TOB ;
    lBoAno       : boolean ;
{$ENDIF}
begin

 result := true ;
{$IFNDEF PGIIMMO}
 lBoAno := true ;
  lBoBDSDyna := GetParamSocSecur('SO_CPBDSDYNA', False) ;
  lBoAnoDyna := (ctxExercice.Suivant.Code <> '') and _GetParamSocSecur('SO_CPANODYNA',false) ;

  if not ( lBoAnoDyna or lBoBDSDyna ) then Exit ;

  lTobListeBDS := nil;

  try

  for i := 0 to TheList.Count - 1 do
  begin

    lFRM := TFRM(TheList.Items[i]^) ;

    if (i = 0) and (lBoBDSDyna) then
     begin // Récup des balances de  situation concernées
      lTobListeBDS := TOB.Create('', nil, -1) ;
      lTobListeBDS.LoadDetailFromSQL('SELECT BSI_CODEBAL FROM CBALSIT WHERE ' +
      'BSI_DATE1 <= "' + UsDateTime(lFRM.DateD) + '" AND ' +
      'BSI_DATE2 >= "' + UsDateTime(lFRM.DateD) + '" AND ' +
      // CA - 18/06/2006 - Il faut mettre à jour uniquement les balances de situation dynamiques.
      'BSI_TYPEBAL="DYN" ORDER BY BSI_CODEBAL') ;
     end ; // if

    if lBoBDSDyna and ( lFRM.fb <> fbAxe1 ) and  (lTobListeBDS.Detail.Count > 0 ) then
     lBoBDSDyna := ExecReqMajBDS( lTobListeBDS, lFRM ) ; // Mise à jour des BDS

    if lBoAnoDyna and ( lFRM.Exo = ctxExercice.EnCours.Code ) then
     begin

      if lBoAno and (lFRM.fb = fbAxe1 ) then
       lBoAno := ExecReqMAJAnoAno(lFRM)
        else
         if lBoAno and ( lFRM.fb = fbAxe2 ) then
          lBoAno := ExecReqMAJBilanAno(lFRM) ;


      if lBoAno and ( lFRM.fb = fbAxe2 ) then
       begin
        lBoAno := ( ExecReqMAJ ( fbGene, false , false , lFRM ) = 1 ) ;

        if lBoAno and ( lFRM.Axe <> '' ) then
         begin
          lFRM.Cpt := lFRM.Axe;
          lFRM.Axe := 'T' ;
          result := ( ExecReqMAJ ( fbAux, false , false , lFRM ) = 1 ) ;
         end ;
       end ; // if
       
     end ;
     
    if not lBoAno and not lBoBDSDyna then
     begin
      result := false ;
      exit ;
     end ;

  end ; // for

 if assigned(lTobListeBDS) then
  FreeAndNil(lTobListeBDS) ; 

 except
  on E : Exception do
   begin
    {$IFNDEF EAGLSERVER}
    Showmessage('Problème lors de l''enregistrement des écritures d''a-nouveaux dynamiques' + #10#13 + E.message) ;
    {$ENDIF}
    V_PGI.IoError := oeSaisie ;
    result := false ;
   end ;
 end ;
{$ENDIF}

end ;

Function ExecReqMAJ ( fb : TFichierBase ; Ano,AvecReseau : boolean ; Var FRM : TFRM ) : LongInt ;
Var SQL : String ;
    lStTable : String ;
BEGIN
  Result := 0;

  if not EstTablePartagee( fbToTable( fb ) ) then
    begin
    // Gestion standard mode PCL OU si la table n'est pas partagée...
    Case fb of
      fbGene : begin
               lStTable := GetTableDossier( FRM.Dossier, 'GENERAUX') ;
               if Not Ano then
                  BEGIN
                  SQL:='UPDATE ' + lStTable + ' SET G_DEBITDERNMVT='+StrFPoint(FRM.Deb)+',  G_CREDITDERNMVT='+StrFPoint(FRM.Cre)+', '
                     +'G_DATEDERNMVT="'+USDateTime(FRM.DateD)+'", G_NUMDERNMVT='+IntToStr(FRM.NumD)+', G_LIGNEDERNMVT='+IntToStr(FRM.LigD)+', '
                     +'G_TOTALDEBIT=G_TOTALDEBIT+'+StrFPoint(FRM.DE+FRM.DS)+', G_TOTALCREDIT=G_TOTALCREDIT+'+StrFPoint(FRM.CE+FRM.CS)+', '
                     +'G_TOTDEBE=G_TOTDEBE+'+StrFPoint(FRM.DE)+', G_TOTCREE=G_TOTCREE+'+StrFPoint(FRM.CE)+', '
                     +'G_TOTDEBS=G_TOTDEBS+'+StrFPoint(FRM.DS)+', G_TOTCRES=G_TOTCRES+'+StrFPoint(FRM.CS)+', '
                     +'G_TOTDEBP=G_TOTDEBP+'+StrFPoint(FRM.DP)+', G_TOTCREP=G_TOTCREP+'+StrFPoint(FRM.CP)+' WHERE G_GENERAL="'+FRM.Cpt+'" ' ;
                  END else
                  BEGIN
                  SQL:='UPDATE ' + lStTable + ' SET G_TOTDEBANO=G_TOTDEBANO+'+StrFPoint(FRM.Deb)+', G_TOTCREANO=G_TOTCREANO+'+StrFPoint(FRM.Cre)+', '
                      +'G_TOTALDEBIT=G_TOTALDEBIT+'+StrFPoint(FRM.Deb)+', G_TOTALCREDIT=G_TOTALCREDIT+'+StrFPoint(FRM.Cre)+', '
                      +'G_TOTDEBE=G_TOTDEBE+'+StrFPoint(FRM.Deb)+', G_TOTCREE=G_TOTCREE+'+StrFPoint(FRM.Cre)+' '
                      +'WHERE G_GENERAL="'+FRM.Cpt+'" ' ;
                  END ;
               end ;
       fbAux : begin
               lStTable := GetTableDossier( FRM.Dossier, 'TIERS') ;
               if Not Ano then
                  BEGIN
                  SQL:='UPDATE ' + lStTable + ' SET T_DEBITDERNMVT='+StrFPoint(FRM.Deb)+',  T_CREDITDERNMVT='+StrFPoint(FRM.Cre)+', '
                     +'T_DATEDERNMVT="'+USDateTime(FRM.DateD)+'", T_NUMDERNMVT='+IntToStr(FRM.NumD)+', T_LIGNEDERNMVT='+IntToStr(FRM.LigD)+', '
                     +'T_TOTALDEBIT=T_TOTALDEBIT+'+StrFPoint(FRM.DE+FRM.DS)+', T_TOTALCREDIT=T_TOTALCREDIT+'+StrFPoint(FRM.CE+FRM.CS)+', '
                     +'T_TOTDEBE=T_TOTDEBE+'+StrFPoint(FRM.DE)+', T_TOTCREE=T_TOTCREE+'+StrFPoint(FRM.CE)+', '
                     +'T_TOTDEBS=T_TOTDEBS+'+StrFPoint(FRM.DS)+', T_TOTCRES=T_TOTCRES+'+StrFPoint(FRM.CS)+', '
                     +'T_TOTDEBP=T_TOTDEBP+'+StrFPoint(FRM.DP)+', T_TOTCREP=T_TOTCREP+'+StrFPoint(FRM.CP)+' WHERE T_AUXILIAIRE="'+FRM.Cpt+'" ' ;
                  END else
                  BEGIN
                  SQL:='UPDATE ' + lStTable + ' SET T_TOTDEBANO=T_TOTDEBANO+'+StrFPoint(FRM.Deb)+', T_TOTCREANO=T_TOTCREANO+'+StrFPoint(FRM.Cre)+', '
                      +'T_TOTALDEBIT=T_TOTALDEBIT+'+StrFPoint(FRM.Deb)+', T_TOTALCREDIT=T_TOTALCREDIT+'+StrFPoint(FRM.Cre)+', '
                      +'T_TOTDEBE=T_TOTDEBE+'+StrFPoint(FRM.Deb)+', T_TOTCREE=T_TOTCREE+'+StrFPoint(FRM.Cre)+' '
                      +'WHERE T_AUXILIAIRE="'+FRM.Cpt+'" ' ;
                  END ;
               end ;
      fbSect : begin
                if VH^.LiaisonY2ViaShare then
                begin
                   lStTable := GetTableDossier( FRM.Dossier, 'CSECTION') ;
                   if Not Ano then
                   BEGIN
                      SQL:='UPDATE ' + lStTable + ' SET CSP_DEBITDERNMVT='+StrFPoint(FRM.Deb)+',  CSP_CREDITDERNMVT='+StrFPoint(FRM.Cre)+', '
                         +'CSP_DATEDERNMVT="'+USDateTime(FRM.DateD)+'", CSP_NUMDERNMVT='+IntToStr(FRM.NumD)+', CSP_LIGNEDERNMVT='+IntToStr(FRM.LigD)+', '
                         +'CSP_TOTALDEBIT=CSP_TOTALDEBIT+'+StrFPoint(FRM.DE+FRM.DS)+', CSP_TOTALCREDIT=CSP_TOTALCREDIT+'+StrFPoint(FRM.CE+FRM.CS)+', '
                         +'CSP_TOTDEBE=CSP_TOTDEBE+'+StrFPoint(FRM.DE)+', CSP_TOTCREE=CSP_TOTCREE+'+StrFPoint(FRM.CE)+', '
                         +'CSP_TOTDEBS=CSP_TOTDEBS+'+StrFPoint(FRM.DS)+', CSP_TOTCRES=CSP_TOTCRES+'+StrFPoint(FRM.CS)+', '
                         +'CSP_TOTDEBP=CSP_TOTDEBP+'+StrFPoint(FRM.DP)+', CSP_TOTCREP=CSP_TOTCREP+'+StrFPoint(FRM.CP)+' '
                         +'WHERE CSP_AXE="'+FRM.Axe+'" AND CSP_SECTION="'+FRM.Cpt+'" ' ;
                   END else
                   BEGIN
                      SQL:='UPDATE ' + lStTable + ' SET S_TOTDEBANO=S_TOTDEBANO+'+StrFPoint(FRM.Deb)+', S_TOTCREANO=S_TOTCREANO+'+StrFPoint(FRM.Cre)+', '
                          +'S_TOTALDEBIT=S_TOTALDEBIT+'+StrFPoint(FRM.Deb)+', S_TOTALCREDIT=S_TOTALCREDIT+'+StrFPoint(FRM.Cre)+', '
                          +'S_TOTDEBE=S_TOTDEBE+'+StrFPoint(FRM.Deb)+', S_TOTCREE=S_TOTCREE+'+StrFPoint(FRM.Cre)+' '
                          +'WHERE S_AXE="'+FRM.Axe+'" AND S_SECTION="'+FRM.Cpt+'" ' ;
                   END ;

                end else
                begin
                   lStTable := GetTableDossier( FRM.Dossier, 'SECTION') ;
                   if Not Ano then
                   BEGIN
                      SQL:='UPDATE ' + lStTable + ' SET S_DEBITDERNMVT='+StrFPoint(FRM.Deb)+',  S_CREDITDERNMVT='+StrFPoint(FRM.Cre)+', '
                         +'S_DATEDERNMVT="'+USDateTime(FRM.DateD)+'", S_NUMDERNMVT='+IntToStr(FRM.NumD)+', S_LIGNEDERNMVT='+IntToStr(FRM.LigD)+', '
                         +'S_TOTALDEBIT=S_TOTALDEBIT+'+StrFPoint(FRM.DE+FRM.DS)+', S_TOTALCREDIT=S_TOTALCREDIT+'+StrFPoint(FRM.CE+FRM.CS)+', '
                         +'S_TOTDEBE=S_TOTDEBE+'+StrFPoint(FRM.DE)+', S_TOTCREE=S_TOTCREE+'+StrFPoint(FRM.CE)+', '
                         +'S_TOTDEBS=S_TOTDEBS+'+StrFPoint(FRM.DS)+', S_TOTCRES=S_TOTCRES+'+StrFPoint(FRM.CS)+', '
                         +'S_TOTDEBP=S_TOTDEBP+'+StrFPoint(FRM.DP)+', S_TOTCREP=S_TOTCREP+'+StrFPoint(FRM.CP)+' '
                         +'WHERE S_AXE="'+FRM.Axe+'" AND S_SECTION="'+FRM.Cpt+'" ' ;
                   END else
                   BEGIN
                      SQL:='UPDATE ' + lStTable + ' SET S_TOTDEBANO=S_TOTDEBANO+'+StrFPoint(FRM.Deb)+', S_TOTCREANO=S_TOTCREANO+'+StrFPoint(FRM.Cre)+', '
                          +'S_TOTALDEBIT=S_TOTALDEBIT+'+StrFPoint(FRM.Deb)+', S_TOTALCREDIT=S_TOTALCREDIT+'+StrFPoint(FRM.Cre)+', '
                          +'S_TOTDEBE=S_TOTDEBE+'+StrFPoint(FRM.Deb)+', S_TOTCREE=S_TOTCREE+'+StrFPoint(FRM.Cre)+' '
                          +'WHERE S_AXE="'+FRM.Axe+'" AND S_SECTION="'+FRM.Cpt+'" ' ;
                   END ;
                end;
               end ;
       fbJal : begin
               lStTable := GetTableDossier( FRM.Dossier, 'JOURNAL') ;
               SQL:='UPDATE ' + lStTable + ' SET J_DEBITDERNMVT='+StrFPoint(FRM.Deb)+',  J_CREDITDERNMVT='+StrFPoint(FRM.Cre)+', '
                   +'J_DATEDERNMVT="'+USDateTime(FRM.DateD)+'", J_NUMDERNMVT='+IntToStr(FRM.NumD)+', '
                   +'J_TOTALDEBIT=J_TOTALDEBIT+'+StrFPoint(FRM.DE+FRM.DS)+', J_TOTALCREDIT=J_TOTALCREDIT+'+StrFPoint(FRM.CE+FRM.CS)+', '
                   +'J_TOTDEBE=J_TOTDEBE+'+StrFPoint(FRM.DE)+', J_TOTCREE=J_TOTCREE+'+StrFPoint(FRM.CE)+','
                   +'J_TOTDEBS=J_TOTDEBS+'+StrFPoint(FRM.DS)+', J_TOTCRES=J_TOTCRES+'+StrFPoint(FRM.CS)+', '
                   +'J_TOTDEBP=J_TOTDEBP+'+StrFPoint(FRM.DP)+', J_TOTCREP=J_TOTCREP+'+StrFPoint(FRM.CP)+' WHERE J_JOURNAL="'+FRM.Cpt+'" ' ;
               end ;
       END ;
    if AvecReseau then SQL:=SQL+FabricReqComp(fb,FRM) ;
    Result := ExecuteSQL(SQL);

    end
  else
  // Gestion des cumuls MULTISOC
    begin
    if CUpdateCumulsMS ( fb, FRM.Cpt, FRM.Axe, FRM.DE+FRM.DS , FRM.CE+FRM.CS , FRM.DP, FRM.CP, FRM.DE, FRM.CE, FRM.DS, FRM.CS, Ano, FRM.Dossier )
      then result := 1 ;
    end ;
END ;

{$IFNDEF EAGLSERVER}
procedure BlocageEntete ( F : TForm ; Treso : Boolean) ;
Var C : TControl ;
BEGIN
C:=TControl(F.FindComponent('E_JOURNAL')) ;
if C<>Nil then
   BEGIN
   TControl(F.FindComponent('E_JOURNAL')).Enabled:=False ;
   TControl(F.FindComponent('E_DEVISE')).Enabled:=False ;
   TControl(F.FindComponent('E_DATECOMPTABLE')).Enabled:=False ;
   C:=TControl(F.FindComponent('E_NATUREPIECE')) ;
   If (C<>NIL) And Treso Then C.Enabled:=False ;
   END ;
END ;

procedure FormatMontant ( GS : THGrid ; ACol,ARow,Decim : integer ) ;
Var X : Double ;
BEGIN
X:=Valeur(GS.Cells[ACol,ARow]) ;
if X=0 then GS.Cells[ACol,ARow]:='' else GS.Cells[ACol,ARow]:=StrS(X,Decim) ;
END ;

procedure VideZone ( GS : THGrid ) ;
Var Col,Lig : integer ;
BEGIN
Col:=GS.Col ; Lig:=GS.Row ;
if ((Lig<=0) or (Lig>=GS.RowCount-1)) then Exit ; if Col<=GS.FixedCols-1 then Exit ;
GS.Cells[Col,Lig]:='' ;
END ;
{$ENDIF}

Function Code2Exige ( Code : String3 ) : TExigeTva ;
BEGIN
Result:=tvaMixte ;
if Code='TD' then Result:=tvaDebit else
 if Code='TE' then Result:=tvaEncais ;
END ;

Function FlagEncais ( Exige : TExigeTva ; SurEnc : boolean ) : String ;
BEGIN
Result:='-' ;
Case Exige of
   tvaDebit  : Result:='-' ;
   tvaEncais : Result:='X' ;
   tvaMixte  : if SurEnc then Result:='X' else Result:='-' ;
   END ;
END ;

Function  Tva2NumBase ( Code : String3 ) : integer ;
Var
{$IFDEF NOVH}
Q   :TQuery;
{$ELSE}
i : integer ;
{$ENDIF}
BEGIN
Result:=0 ; if Code='' then Exit ;
{$IFDEF NOVH}
Q:=OpenSQL('Select CC_CODE, CC_LIBRE from CHOIXCOD Where CC_TYPE="TX1" AND CC_LIBRE>"0" AND CC_LIBRE<="4" AND CC_CODE="'+
Code +'"',True,-1,'',true) ;
if Not Q.EOF  and (Q.FindField('CC_LIBRE').asstring <> '') then
          Result := StrToInt(Q.FindField('CC_LIBRE').asstring);
Ferme(Q);
{$ELSE}
for i:=1 to 4 do if VH^.NumCodeBase[i]=Code then BEGIN Result:=i ; Exit ; END ;
{$ENDIF}
END ;



procedure AppelRequete(Var Q : TQuery ; What : Integer; Cpte, Axe : string);
Var St : String ;
BEGIN
Q:=Nil ;
Case What Of
  0 : St:='SELECT G_NATUREGENE,G_REGIMETVA,G_DATEDERNMVT,G_NUMDERNMVT, '+
          'G_LIGNEDERNMVT,G_DEBITDERNMVT,G_CREDITDERNMVT,G_TOTALDEBIT,G_TOTALCREDIT, '+
          'G_TOTDEBE,G_TOTCREE,G_TOTDEBS,G_TOTCRES,G_TOTDEBANO,G_TOTCREANO,G_VENTILABLE, '+
          'G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5, G_COLLECTIF, '+
          'G_LETTRABLE,G_POINTABLE,G_TVA,G_TPF,G_TVAENCAISSEMENT,G_TVASURENCAISS FROM GENERAUX '+
          'WHERE G_GENERAL="'+Cpte+'"' ;
  1 : St:='SELECT T_NATUREAUXI,T_REGIMETVA,T_DATEDERNMVT,T_NUMDERNMVT,T_COLLECTIF, '+
          'T_LIGNEDERNMVT,T_DEBITDERNMVT,T_CREDITDERNMVT,T_TOTALDEBIT,T_TOTALCREDIT, '+
          'T_TOTDEBE,T_TOTCREE,T_TOTDEBS,T_TOTCRES,T_TOTDEBANO,T_TOTCREANO,T_LETTRABLE, '+
          'T_TVAENCAISSEMENT,T_PAYEUR,T_ISPAYEUR, T_AUXILIAIRE FROM TIERS '+
          'WHERE T_AUXILIAIRE="'+Cpte+'"' ;
  2 : St:='SELECT S_SECTION,S_AXE,S_DATEDERNMVT, '+
          'S_NUMDERNMVT,S_LIGNEDERNMVT,S_DEBITDERNMVT,S_CREDITDERNMVT,S_TOTALDEBIT,S_TOTALCREDIT, '+
          'S_TOTDEBE,S_TOTCREE,S_TOTDEBS,S_TOTCRES,S_TOTDEBANO,S_TOTCREANO FROM SECTION '+
          'WHERE S_SECTION="'+Cpte+'" AND S_AXE="'+AXE+'"' ;
  3 : St:='SELECT J_JOURNAL,J_NATUREJAL,J_COMPTEURNORMAL,J_COMPTEURSIMUL,J_DATEDERNMVT,'+
          'J_NUMDERNMVT,J_DEBITDERNMVT,J_CREDITDERNMVT,J_TOTALDEBIT,J_TOTALCREDIT,'+
          'J_TOTDEBE,J_TOTCREE,J_TOTDEBS,J_TOTCRES,J_CONTREPARTIE,J_MODESAISIE FROM JOURNAL '+
          'WHERE J_JOURNAL="'+Cpte+'"' ;
  4 : St:='SELECT BJ_BUDJAL,BJ_NATJAL,BJ_COMPTEURNORMAL,BJ_COMPTEURSIMUL '+
          'FROM BUDJAL WHERE BJ_BUDJAL="'+Cpte+'"' ;
  5 : St:='SELECT CR_TYPE,CR_CORRESP,CR_LIBELLE,CR_LIBRETEXTE1 from CORRESP WHERE CR_TYPE=:CRTYPE AND CR_CORRESP=:CRCORRESP' ;
  END ;
Q:=OpenSql(St,TRUE) ;
END ;

{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
Function AlimLTabCptLu(What : Integer ; Q : TQuery ; L : HTStringList ; ListeCptFaux : TList ; Var CptLu : TCptLu ) : Boolean ;
Var i : Integer ;
    Cpt,MessSup,CptAssocie,Axe : String ;
    OkOk,Existe,RechercheAFaire : Boolean ;
    X : DelInfo ;
    LeCptLuP : TCptLuP ;
{  0 : Généraux 1 : Tiers 2 : Section 3 : Journaux 4 : Journaux Budgétaires }
BEGIN
Result:=FALSE ;
Cpt:=Trim(CptLu.Cpt) ; If Cpt='' Then Exit ;
OkOk:=FALSE ; Existe:=FALSE ; CptAssocie:=CptLu.Collectif ; Axe:=Trim(CptLu.Axe) ;
i:=L.IndexOf(Cpt) ;
if i>-1 then
   BEGIN
   LeCptLuP:=TCptLuP(L.Objects[i]) ; OkOk:=TRUE ; Existe:=LeCptLuP.T.Ok ; If Existe Then Result:=TRUE ;
   ClassToRec(CptLu,LeCptLuP) ;
   END Else
   BEGIN
   LeCptLuP:=TCptLuP.Create ;(*Else LeCptLu.Cpt:=Cpt ;*)
   LeCptLuP.T.Axe:='-----' ;
   LeCptLuP.T.Cpt:=Cpt ;
   END ;
//AppelRequete(Q, What, Cpt, Axe); // CA - 14/11/2007 - déplacé. Inutile ici
// GP pas de chance : q = nil en recalcul tva encaissement !
//RechercheAFaire:=((Not OkOk) Or (Not Existe)) And (Q<>NIL) ;
RechercheAFaire:=((Not OkOk) Or (Not Existe))  ;
If RechercheAFaire Then
   BEGIN
   If What=2 Then
     If Axe='' Then Axe:='A1' ;
   AppelRequete(Q, What, Cpt, Axe); // CA - 14/11/2007

   if Q.Eof then
      BEGIN
      If (ListeCptFaux<>NIL) And (What in [0,1,2]) Then
         BEGIN
         X:=DelInfo.Create ; X.Gen:='' ; X.aux:='' ; X.Sect:='' ;
         X.GenACreer:=FALSE ; X.AuxACreer:=FALSE ; X.SectACreer:=FALSE ;
         X.Axe:='' ;
         Case What Of
            0 : BEGIN
                X.Gen:=Cpt ; X.GenACreer:=TRUE ;
                END ;
            1 : BEGIN
                X.Aux:=Cpt ; X.AuxACreer:=TRUE ;
                END ;
            2 : BEGIN
                X.Sect:=Cpt ; X.SectACreer:=TRUE ; X.Axe:=Axe ;
                END ;
            END ;
         MessSup:=MessErreurCompte(What,CptLu) ;
         X.LeCod:=Cpt ; X.LeLib:=MessSup ;
         ListeCptFaux.Add(X);
         END ;
      END Else
      BEGIN
      LeCptLuP.T.Ok:=TRUE ; Result:=TRUE ;
      Case What Of
        0 : BEGIN
            LeCptLuP.T.Nature:=Trim(Q.FindField('G_NATUREGENE').AsString) ;
            LeCptLuP.T.Pointable:=Q.FindField('G_POINTABLE').AsString='X' ;
            LeCptLuP.T.Lettrable:=Q.FindField('G_LETTRABLE').AsString='X' ;
            LeCptLuP.T.EstColl:=Q.FindField('G_COLLECTIF').AsString='X' ;
            LeCptLuP.T.RegimeTva:=Q.FindField('G_REGIMETVA').AsString ;
            LeCptLuP.T.Ventilable:=(Q.FindField('G_VENTILABLE').AsString='X') ;
            if Q.FindField('G_VENTILABLE1').AsString='X' then LeCptLuP.T.Axe[1]:='X' ;
            if Q.FindField('G_VENTILABLE2').AsString='X' then LeCptLuP.T.Axe[2]:='X' ;
            if Q.FindField('G_VENTILABLE3').AsString='X' then LeCptLuP.T.Axe[3]:='X' ;
            if Q.FindField('G_VENTILABLE4').AsString='X' then LeCptLuP.T.Axe[4]:='X' ;
            if Q.FindField('G_VENTILABLE5').AsString='X' then LeCptLuP.T.Axe[5]:='X' ;
            LeCptLuP.T.Tva:=Q.FindField('G_TVA').AsString ;
            LeCptLuP.T.Tpf:=Q.FindField('G_TPF').AsString ;
            LeCptLuP.T.TvaEnc:=Trim(Q.FindField('G_TVAENCAISSEMENT').AsString) ;
            LeCptLuP.T.TvaEncHT:=Trim(Q.FindField('G_TVASURENCAISS').AsString) ;
            END ;
        1 : BEGIN
            LeCptLuP.T.Nature:=Trim(Q.FindField('T_NATUREAUXI').AsString) ;
            LeCptLuP.T.CollTiers:=Trim(Q.FindField('T_COLLECTIF').AsString) ;
            LeCptLuP.T.RegimeTva:=Trim(Q.FindField('T_REGIMETVA').AsString) ;
            LeCptLuP.T.Lettrable:=Q.FindField('T_LETTRABLE').AsString='X' ;
            LeCptLuP.T.TvaEnc:=Trim(Q.FindField('T_TVAENCAISSEMENT').AsString) ;
            LeCptLuP.T.TP:=Trim(Q.FindField('T_PAYEUR').AsString) ;
            LeCptLuP.T.ISTP:=Trim(Q.FindField('T_ISPAYEUR').AsString)='X' ;
            END ;
        2 : BEGIN
            LeCptLuP.T.Axe:=Trim(Q.FindField('S_AXE').AsString) ;
            END ;
        3 : BEGIN
            LeCptLuP.T.Nature:=Trim(Q.FindField('J_NATUREJAL').AsString) ;
            LeCptLuP.T.SoucheN:=Q.FindField('J_COMPTEURNORMAL').AsString ;
            LeCptLuP.T.SoucheS:=Q.FindField('J_COMPTEURSIMUL').AsString ;
            LeCptLuP.T.ModeSaisie:=Q.FindField('J_MODESAISIE').AsString ;
            LeCptLuP.T.CollTiers:=Trim(Q.FindField('J_CONTREPARTIE').AsString) ;
            END ;
        4 : BEGIN
            LeCptLuP.T.Nature:=Q.FindField('BJ_NATJAL').AsString ;
            LeCptLuP.T.SoucheN:=Q.FindField('BJ_COMPTEURNORMAL').AsString ;
            LeCptLuP.T.SoucheS:=Q.FindField('BJ_COMPTEURSIMUL').AsString ;
            END ;
        END ;
      END ;
   Ferme(Q) ;
   END ;
If CptAssocie<>'' Then LeCptLuP.T.Collectif:=CptAssocie ;
If Not OkOk Then L.AddObject(Cpt,LeCptLuP) ;
ClassToRec(CptLu,LeCptLuP) ;
END ;

Procedure ClassToRec(Var C1 : TCptLu ; C2 : TCptLuP) ;
BEGIN
With C1,C2 Do
  BEGIN
  Ok                   :=T.Ok          ; Cpt       :=T.Cpt;
  Collectif            :=T.Collectif   ; CollTiers :=T.CollTiers;
  Ventilable           :=T.Ventilable  ; VentilableDansFichier:=T.VentilableDansFichier ;
  Nature               :=T.Nature      ; RegimeTva :=T.RegimeTva;
  Pointable            :=T.Pointable   ; Lettrable :=T.Lettrable  ;
  SoucheN              :=T.SoucheN     ; SoucheS   :=T.SoucheS ;
  Axe                  :=T.Axe         ; LastDate  :=T.LastDate ;
  LastNum              :=T.LastNum     ; LastNumL  :=T.LastNumL;
  DebitMvt             :=T.DebitMvt    ; CreditMvt :=T.CreditMvt ;
  TotDN                :=T.TotDN       ; TotCN     :=T.TotCN ;
  TotDN1               :=T.TotDN1      ; TotCN1    :=T.TotCN1 ;
  TotDAno              :=T.TotDAno     ; TotCAno   :=T.TotCAno ;
  Tva                  :=T.Tva         ; TPF       :=T.TPF ;
  OkMajSolde           :=T.OkMajSolde  ; TvaEnc    :=T.TvaEnc ;
  TvaEncHT             :=T.TvaEncHT    ; TP        :=T.TP ;
  IsTP                 :=T.IsTP        ; ModeSaisie:=T.ModeSaisie ;
  QualifPiece          :=T.QualifPiece ; EstColl   :=T.EstColl ;
  END ;
END ;

Function MessErreurCompte(What : Integer ; CptLu : TCptLu) : String ;
Var Racine : String ;

BEGIN
Result:='' ;
Case What Of
  0 : Result:='Le compte général n''existe pas.' ;
  1 : Result:='Le compte auxiliaire n''existe pas.' ;
  2 : Result:='La section analytique n''existe pas.' ;
  END ;

Case What Of
  0 : BEGIN
      Racine:=Copy(CptLu.Cpt,1,1) ;
      If Racine='6' Then  Result:=Result+' Compte de charge' Else
         If Racine='7' Then  Result:=Result+' Compte de produit' ;
      If CptLu.VentilableDansFichier Then Result:=Result+' Ventilable' ;
      END ;
  1 : BEGIN
      Racine:=Copy(CptLu.Collectif,1,2) ;
      If Racine='40' Then Result:=Result+' Fournisseur' Else
         If Racine='41' Then Result:=Result+' Client' ;
      If CptLu.Collectif<>'' Then Result:=Result+' - '+CptLu.Collectif ;
      END ;
  2 : Result:=Result+' Axe '+CptLu.Axe ;
  END ;
Result:=TraduireMemoire(Result) ;
END ;

Function ChercheCptLu(L : HTStringList ; Var CptLu : TCptLu) : Boolean ;
Var LeCptLuP : TCptLuP ;
    i : Integer ;
    Cpt : String ;
BEGIN
Result:=FALSE ; Cpt:=Trim(CptLu.Cpt) ; If L=NIL Then Exit ;
i:=L.IndexOf(Cpt) ;
if i>-1 then
  BEGIN
  If L.Objects[i]<>NIL Then BEGIN
  LeCptLuP:=TCptLuP(L.Objects[i]) ;
  ClassToRec(CptLu,LeCptLuP) ;
  END ;
  Result:=TRUE ;
  END ;
END ;
{$ENDIF}
{$ENDIF}

Procedure ElementsTvaEnc ( M : RMVT ; ForceFlag : boolean ; ExiConnue : String = ''; AccepteOD : boolean = false ) ;

  procedure _VideListe(L: HTStringList);
  var
    i: integer;
  begin
    if L = nil then Exit;
    for i:=0 to L.Count-1 do
      if TObject(L.Objects[i])<>Nil then TObject(L.Objects[i]).Free ;
    L.Clear ;
  end;

{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
 Var InfoImp : TInfoImport ; QFiche : TQFiche;
{$ENDIF}
{$ENDIF}
begin
{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
  FillChar(InfoImp,SizeOf(InfoImp),#0) ;
  InfoImp.LGenLu:=HTStringList.Create ;
  InfoImp.LAuxLu:=HTStringList.Create ;
  InfoImp.LJalLu:=HTStringList.Create ;
  ElementsTvaEncRevise (M,  ForceFlag, InfoImp, QFiche, ExiConnue, AccepteOD) ;
  _VideListe(InfoImp.LGenLu);
  FreeandNil(InfoImp.LGenLu) ;
  _VideListe(InfoImp.LAuxLu);
  FreeandNil(InfoImp.LAuxLu) ;
  _VideListe(InfoImp.LJalLu);
  FreeandNil(InfoImp.LJalLu) ;
  FreeandNil(InfoImp) ;
{$ENDIF}
{$ENDIF}
end;

// ME 04/06/2007 unification de la fonction  ElementsTvaEnc avec  TraitementSurEcr fiche 10537
{$IFNDEF CCADM}
{$IFNDEF PGIIMMO}
Procedure ElementsTvaEncRevise ( M : RMVT ; ForceFlag : boolean ; Var InfoImp : TInfoImport ; QFiche : TQFiche; ExiConnue : String = ''; AccepteOD : boolean = false ) ;
Var TPiece : TList ;
    Q : TQuery ;
    O : TOB ;
    NbTiers,Lig,LigTiers,NumBase,i : integer ;
    SorteTva  : TSorteTva ;
    ExigeTva  : TExigeTva ;
    Solde,TTC,Coef : double ;
    Okok,OkTiers,ExisteEnc : boolean ;
    NaturePiece,NatureJal,RegimeTva,RegimeEnc,StE,NatGene,CodeTva : String3 ;
    Auxi,Gene, Coll : String ;
    TabTvaEnc : Array[1..5] of Double ;
    CptLu,CptLuJ : TCptLu ;
    EncaisseChange : Boolean ;
BEGIN
{#TVAENC}
{$IFDEF NOVH}
if not GetParamSocSecur('SO_OUITVAENC',False) then exit;
{$ELSE}
if Not VH^.OuiTvaEnc then Exit ;
{$ENDIF}

if V_PGI.IoError<>oeOk then Exit ;
if M.Jal='' then Exit ;
Okok:=True ; ExisteEnc:=False ;
{Identification nature facture}
NaturePiece:=M.Nature ;
if ((NATUREPIECE<>'FC') and (NATUREPIECE<>'AC') and (NATUREPIECE<>'FF') and (NATUREPIECE<>'AF') and
   ((AccepteOD) and (NATUREPIECE<>'OD'))) then Exit ;
{Identification journal facture}
NatureJal:='' ; TTC:=0 ;
CptLuJ.Cpt:=M.Jal ;
If AlimLTabCptLu(3,QFiche[3],InfoImp.LJalLu,NIL,CptLuJ) Then NatureJal:=CptLuJ.Nature ;
SorteTva:=stvDivers ;
if NatureJal='VTE' then SorteTva:=stvVente else if NatureJal='ACH' then SorteTva:=stvAchat ;
if SorteTva=stvDivers then Exit ;
{Constitution d'une liste des lignes de l'écriture}
TPiece:=TList.Create ;
Q:=OpenSQL('Select * from ECRITURE Where '+WhereEcriture(tsGene,M,False),True,-1,'',true) ;
While Not Q.EOF do
   BEGIN
   CptLu.Cpt:=Q.FindField('E_GENERAL').AsString ; AlimLTabCptLu(0,QFiche[0],InfoImp.LGenLu,NIL,CptLu) ;
   CptLu.Cpt:=Q.FindField('E_AUXILIAIRE').AsString ; AlimLTabCptLu(1,QFiche[1],InfoImp.LAuxLu,NIL,CptLu) ;
   O:=TOB.Create('ECRITURE',Nil,-1) ;
   O.SelectDB('',Q,True) ; TPiece.Add(O) ;
   Q.Next ;
   END ;
Ferme(Q) ;
{Détermination du nombre de tiers}
NbTiers:=0 ; LigTiers:=-1 ;
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOB(TPiece[Lig]) ; OkTiers:=False ;
    if ((O.GetValue('E_ECHE')='X') and (O.GetValue('E_NUMECHE')>=1)) and (O.GetString('E_TYPEMVT')='TTC') then
       BEGIN
       if O.GetValue('E_NUMECHE')=1 then OkTiers:=True ;
       TTC:=Arrondi(TTC+O.GetValue('E_DEBIT')-O.GetValue('E_CREDIT'),V_PGI.OkDecV) ;
       END ;
    if OkTiers then
       BEGIN
       Inc(NbTiers) ;
       if NbTiers=1 then LigTiers:=Lig else BEGIN Okok:=False ; Break ; END ;
       END ;
    END ;
if ((Not Okok) or (NbTiers<>1) or (LigTiers<0) or (TTC=0)) then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Lecture paramètres tiers}
O:=TOB(TPiece[LigTiers]) ;
Coll:=O.GetValue('E_GENERAL') ; if Not EstCollFact(Coll) then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
RegimeTva:='' ; RegimeEnc:='' ;
if O.GetValue('E_AUXILIAIRE')<>'' then
   BEGIN
   Auxi:=O.GetValue('E_AUXILIAIRE') ;
   CptLu.Cpt:=Auxi ;
   If ChercheCptLu(InfoImp.LAuxLu,CptLu) Then BEGIN RegimeTva:=CptLu.RegimeTva ; RegimeEnc:=CptLu.TvaEnc ; END ;
   END else
   BEGIN
   Gene:=O.GetValue('E_GENERAL') ;
   CptLu.Cpt:=Gene ;
   If ChercheCptLu(InfoImp.LGenLu,CptLu) Then BEGIN RegimeTva:=CptLu.RegimeTva ; RegimeEnc:=CptLu.TvaEnc ; END ;
   END ;
if ExiConnue<>'' then RegimeEnc:=ExiConnue ;
if ((RegimeTva='') or ((SorteTva=stvAchat) and (RegimeEnc=''))) then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Affectation exigibilité Tva}
if SorteTva=stvAchat then ExigeTva:=Code2Exige(RegimeEnc) else ExigeTva:=tvaMixte ;
{Si Fournisseur débit --> débit donc rien à traiter}
if ExigeTva=tvaDebit then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Affectation des régimes, des tva,tpf et enc O/N sur les lignes}
EncaisseChange:=FALSE ;
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOB(TPiece[Lig]) ; Gene:=O.GetValue('E_GENERAL') ;
    CptLu.Cpt:=Gene ;
    If ChercheCptLu(InfoImp.LGenLu,CptLu) Then
      BEGIN
      NatGene:=CptLu.Nature ;
      if ((NatGene='CHA') or (NatGene='PRO') or (NatGene='IMO')) then
         BEGIN
         If ((O.GetValue('E_TVA')='') Or ForceFlag) Then O.PutValue('E_TVA',CptLu.Tva) ;
         If ((O.GetValue('E_TPF')='') Or ForceFlag) Then O.PutValue('E_TPF',CptLu.Tpf) ;
         If (O.GetValue('E_TVAENCAISSEMENT')='') Or ForceFlag Then
           BEGIN
           StE:=FlagEncais(ExigeTva,(CptLu.TvaEncHT='X')) ;
           END else StE:=O.GetValue('E_TVAENCAISSEMENT') ;
         if StE='X' then ExisteEnc:=True ;
          if O.Getvalue('E_REGIMETVA') = '' then
         O.PutValue('E_REGIMETVA',RegimeTva) ;
         If (O.GetValue('E_TVAENCAISSEMENT')<>StE) Then
           begin
           O.PutValue('E_TVAENCAISSEMENT',StE) ;
           EncaisseChange:=TRUE ;
           end ;
         END else
           O.PutValue('E_TVA','') ;
      END ;
    END ;
{Si aucune ligne en encaissement --> aucun traitement}
if (Not ExisteEnc) And (Not EncaisseChange) then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Calcul du tableau des bases Tva}
FillChar(TabTvaEnc,Sizeof(TabTvaEnc),#0) ;
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOB(TPiece[Lig]) ;
    CodeTva:=O.GetValue('E_TVA') ; if CodeTva='' then Continue ;
    if SorteTva=stvVente then
      Solde:=O.GetValue('E_CREDIT')-O.GetValue('E_DEBIT')
      else
      Solde:=O.GetValue('E_DEBIT')-O.GetValue('E_CREDIT') ;
    if O.GetValue('E_TVAENCAISSEMENT')='X' then
     BEGIN
       NumBase:=Tva2NumBase(CodeTva) ;
       if NumBase>0 then
         TabTvaEnc[NumBase]:=TabTvaEnc[NumBase]+Solde ;
     END else
     BEGIN
       TabTvaEnc[5]:=TabTvaEnc[5]+Solde ;
     END ;
    END ;
{Report du tableau des bases sur les échéances}
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOB(TPiece[Lig]) ;
    if ((O.GetValue('E_ECHE')='X') and (O.GetValue('E_NUMECHE')>0)) and (O.GetString('E_TYPEMVT')='TTC') then
       BEGIN
       Solde:=O.GetValue('E_DEBIT')-O.GetValue('E_CREDIT') ;
       Coef:=Solde/TTC ;
       for i:=1 to 4 do O.PutValue('E_ECHEENC'+IntToStr(i),Arrondi(TabTvaEnc[i]*Coef,V_PGI.OkDecV)) ;
       O.PutValue('E_ECHEDEBIT',Arrondi(TabTvaEnc[5]*Coef,V_PGI.OkDecV)) ;
       O.PutValue('E_EMETTEURTVA','X') ;
       END ;
    END ;
{Mise à jour fichier}
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOB(TPiece[Lig]) ;
    if Not O.IsOneModifie then Continue ;
    O.SetAllModifie(True) ;
    if Not O.UpDateDB(False) then BEGIN V_PGI.IoError:=oeUnknown ; Break ; END ;
    END ;
{Dispose mémoire}
VideListe(TPiece) ; TPiece.Free ;
END ;
{$ENDIF}
{$ENDIF}

function EncodeLC ( R : RMVT ) : String ;
BEGIN
Result:=R.Jal+';'+R.Exo+';'+DateToStr(R.DateC)+';'+IntToStr(R.Num)+';'
       +IntToStr(R.NumLigne)+';'+IntToStr(R.NumEche)+';'+R.Simul+';'+R.Nature+';' ;
END ;

function DecodeLC ( StM : String ) : RMVT ;
Var St,StC : String ;
    R      : RMVT ;
BEGIN
FillChar(R,Sizeof(R),#0) ; St:=StM ;
StC:=Trim(ReadTokenSt(StM)) ; R.Jal:=StC ;
StC:=Trim(ReadTokenSt(StM)) ; R.Exo:=StC ;
StC:=Trim(ReadTokenSt(StM)) ; R.DateC:=StrToDate(StC) ;
StC:=Trim(ReadTokenSt(StM)) ; R.Num:=StrToInt(StC) ;
StC:=Trim(ReadTokenSt(StM)) ; R.NumLigne:=StrToInt(StC) ;
StC:=Trim(ReadTokenSt(StM)) ; R.NumEche:=StrToInt(StC) ;
StC:=Trim(ReadTokenSt(StM)) ; R.Simul:=StC ;
StC:=Trim(ReadTokenSt(StM)) ; R.Nature:=StC ;
Result:=R ;
END ;

{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}

{$IFDEF EAGLCLIENT}
{$ELSE}
{procedure AttribParams ( QX : TQuery ; D,C : Double ; T : TTypeExo ) ;
BEGIN
QX.ParamByName('DD').AsFloat:=D ; QX.ParamByName('CD').AsFloat:=C ;
Case T of
   teEncours : BEGIN
               QX.ParamByName('DE').AsFloat:=D ; QX.ParamByName('CE').AsFloat:=C ;
               QX.ParamByName('DS').AsFloat:=0 ; QX.ParamByName('CS').AsFloat:=0 ;
               QX.ParamByName('DP').AsFloat:=0 ; QX.ParamByName('CP').AsFloat:=0 ;
               END ;
   teSuivant : BEGIN
               QX.ParamByName('DE').AsFloat:=0 ; QX.ParamByName('CE').AsFloat:=0 ;
               QX.ParamByName('DS').AsFloat:=D ; QX.ParamByName('CS').AsFloat:=C ;
               QX.ParamByName('DP').AsFloat:=0 ; QX.ParamByName('CP').AsFloat:=0 ;
               END ;
 tePrecedent : BEGIN
               QX.ParamByName('DE').AsFloat:=0 ; QX.ParamByName('CE').AsFloat:=0 ;
               QX.ParamByName('DS').AsFloat:=0 ; QX.ParamByName('CS').AsFloat:=0 ;
               QX.ParamByName('DP').AsFloat:=D ; QX.ParamByName('CP').AsFloat:=C ;
               END ;
   END ;
END ;}
{$ENDIF}

{$IFNDEF EAGLSERVER}
procedure ZeroBlanc ( P : TPanel ) ;
Var i : integer ;
    C : TControl ;
BEGIN
for i:=0 to P.ControlCount-1 do
    BEGIN
    C:=P.Controls[i] ;
    if C.ClassType=THNumEdit then THNumEdit(C).Value:=0 else
    if ((C.ClassType=TEdit) and (C.Name<>'NONCALC')) then TEdit(C).Text:='' ;
    END ;
END ;
{$ENDIF}

Function StrToTypCtr(St : String3) : ttCtrPtr ;
begin
Result:=Auchoix ;
If St='LIG' then Result:=Ligne else
   if St='PID' then Result:=PiedDC else
      if St='PIS' then Result:=PiedSolde else
         if St='MAN' then Result:=AuChoix ;
end ;

{======================== OBJECT JOURNAL ====================================}
{***********A.G.L.***********************************************
Auteur  ...... : Piot
Créé le ...... : 22/02/2005
Modifié le ... :   /  /
Description .. : - 22/02/2005 - LG - Ajout de l'accelerateur de saisie
Mots clefs ... : 
*****************************************************************}
Function TSAJAL.SelectJAL( Jal : String ) : String ;
BEGIN
SelectJAL:='Select J_JOURNAL, J_NATUREJAL, J_MULTIDEVISE, J_COMPTEURNORMAL, J_COMPTEURSIMUL, J_COMPTEINTERDIT, '
          +'J_LIBELLE, J_COMPTEAUTOMAT, J_ABREGE, J_TYPECONTREPARTIE, J_CONTREPARTIE, J_AXE, '
          +'J_VALIDEEN, J_VALIDEEN1, J_FERME, J_MODESAISIE, J_EFFET,J_ACCELERATEUR From JOURNAL Where J_JOURNAL="'+Jal+'"' ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Piot
Créé le ...... : 22/02/2005
Modifié le ... :   /  /    
Description .. : - LG - 22/02/2005 - Ajout de l'accelerateur de saisie
Mots clefs ... : 
*****************************************************************}
procedure TSAJAL.QueryToJal ( JQ : TQuery ) ;
Var St : String ;
BEGIN
JOURNAL:=JQ.FindField('J_JOURNAL').AsString               ; NATUREJAL:=JQ.FindField('J_NATUREJAL').AsString ;
MULTIDEVISE:=(JQ.FindField('J_MULTIDEVISE').AsString='X') ;
COMPTEURNORMAL:=JQ.FindField('J_COMPTEURNORMAL').AsString ; COMPTEURSIMUL:=JQ.FindField('J_COMPTEURSIMUL').AsString ;
COMPTEINTERDIT:=JQ.FindField('J_COMPTEINTERDIT').AsString ; COMPTEAUTOMAT:=JQ.FindField('J_COMPTEAUTOMAT').AsString ;
LIBELLEJOURNAL:=JQ.FindField('J_LIBELLE').AsString        ; ABREGEJOURNAL:=JQ.FindField('J_ABREGE').AsString ;
AXE:=JQ.FindField('J_AXE').AsString                       ; OKFERME:=JQ.FindField('J_FERME').AsString='X' ;
VALIDEEN:=JQ.FindField('J_VALIDEEN').AsString             ; VALIDEEN1:=JQ.FindField('J_VALIDEEN1').AsString ;
ModeSaisieJal:=JQ.FindField('J_MODESAISIE').AsString      ;
JALEFFET:=JQ.FindField('J_EFFET').AsString='X'            ;
TRESO.Cpt:=JQ.FindField('J_CONTREPARTIE').AsString        ; TRESO.STypCtr:=JQ.FindField('J_TYPECONTREPARTIE').AsString ;
TRESO.TypCtr:=StrToTypCtr(TRESO.STypCtr) ;
If SaisieTreso And (Not Effet) Then COMPTEAUTOMAT:='' ;
St:=COMPTEAUTOMAT ; NbAuto:=0 ; While ReadTokenSt(St)<>'' do Inc(NbAuto) ;
if ((NATUREJAL='BQE') or (NATUREJAL='CAI') or (JALEFFET)) And (Not SaisieTreso) then BEGIN COMPTEAUTOMAT:=COMPTEAUTOMAT+TRESO.Cpt+';' ; Inc(NbAuto) ; END ;
If SaisieTreso And (Not Effet) Then ChargeCompteTreso ;
J_ACCELERATEUR := JQ.FindField('J_ACCELERATEUR').AsString ;
END ;

{---------------------------------------------------------------------------------------}
Constructor TSAJAL.Create( Jal : String  ; SaisTreso : Boolean) ;
{---------------------------------------------------------------------------------------}
Var JQ : TQuery ;
begin
  inherited Create ;
  SaisieTreso:=SaisTreso ; Effet:=FALSE ;
  DateRef := V_PGI.DateEntree;{JP 17/01/08 : FQ 18563}
  JQ:=OpenSQL(SelectJAL(Jal),True,-1,'',true) ;
  if Not JQ.EOF then QueryToJal(JQ) ;
  Ferme(JQ) ;
end ;

{---------------------------------------------------------------------------------------}
Constructor TSAJAL.CreateEff( Jal : String ) ;
{---------------------------------------------------------------------------------------}
Var JQ : TQuery ;
begin
  Inherited Create ;
  SaisieTreso:=TRUE ; Effet:=TRUE ;
  DateRef := V_PGI.DateEntree;{JP 17/01/08 : FQ 18563}
  JQ:=OpenSQL(SelectJAL(Jal),True,-1,'',true) ;
  if Not JQ.EOF then QueryToJal(JQ) ;
  Ferme(JQ) ;
end ;

{---------------------------------------------------------------------------------------}
procedure TSAJAL.SetDateExo(Value : TDateTime);
{---------------------------------------------------------------------------------------}
var
  lStCodeExo : string;
begin
  if Value <> FDateRef then FDateRef := Value;
  lStCodeExo := QuelExoDt(Value);
       if lStCodeExo = GetEncours.Code then FTypeExo := teEncours
  else if lStCodeExo = GetSuivant.Code then FTypeExo := teSuivant
                                       else FTypeExo := tePrecedent;
end;

{---------------------------------------------------------------------------------------}
procedure TSAJAL.PutDate(DateC : string);
{---------------------------------------------------------------------------------------}
begin
  if IsValidDate(DateC) then DateRef := StrToDate(DateC)
                        else DateRef := V_PGI.DateEntree;
end;

{JP 17/01/08 : FQ 18563 : Gestion du type d'exercice pour le calcul des Totaux
{---------------------------------------------------------------------------------------}
Function TSAJAL.ChargeCompteTreso : Boolean ;
{---------------------------------------------------------------------------------------}
var
  Q         : TQuery;
  SQL       : string;
  OkOk      : Boolean;
  lStLettre : string;{JP 17/01/08 : FQ 18563}
  ChpCredit : string;{JP 17/01/08 : FQ 18563}
  ChpDebit  : string;{JP 17/01/08 : FQ 18563}
begin
  OkOk := False;

  case FTypeExo of
    teEncours   : lStLettre := 'E';
    teSuivant   : lStLettre := 'S';
    tePrecedent : lStLettre := 'P';
  end;

  ChpCredit := 'G_TOTCRE' + lStLettre;
  ChpDebit  := 'G_TOTDEB' + lStLettre;

  // Gestion des cumuls MULTISOC
  if EstTablePartagee( 'GENERAUX' )
    then SQL:='Select G_GENERAL, ' + ChpDebit + ', ' + ChpCredit + ', G_VENTILABLE, G_COLLECTIF FROM GENERAUXMS WHERE G_GENERAL="'+TRESO.Cpt+'"'
    else SQL:='Select G_GENERAL, ' + ChpDebit + ', ' + ChpCredit + ', G_VENTILABLE, G_COLLECTIF FROM GENERAUX WHERE G_GENERAL="'+TRESO.Cpt+'"' ;
  Q:=OpenSQL(SQL,TRUE,-1,'',true) ;
  if Not Q.Eof then begin
    TRESO.Ventilable := Q.FindField('G_VENTILABLE').AsString='X' ;
    TRESO.TotD       := Q.FindField(ChpDebit ).AsFloat ;{JP 17/01/08 : FQ 18563}
    TRESO.TotC       := Q.FindField(ChpCredit).AsFloat ;{JP 17/01/08 : FQ 18563}
    TRESO.Collectif  := Q.FindField('G_COLLECTIF').AsString='X' ;
    OkOk:=TRUE ;
  end ;
Ferme(Q) ;
If OkOk Then
   BEGIN
   If Not Effet Then TRESO.DevBqe:=GetDeviseBanque(Treso.Cpt) ;
   END ;
Result:=OkOk ;
END ;

{$IFNDEF EAGLSERVER}
procedure ProrateQteAnal ( GS : THGrid ; Lig : integer ; OQ,NQ : Double ; Ind : integer ) ;
Var OBA  : TObjAna ;
    O    : TOBM ;
BEGIN
OBA:=TObjAna(GS.Objects[SA_DateC,Lig]) ; if OBA=Nil then Exit ;
O:=GetO(GS,Lig) ; if O=Nil then Exit ;
ProrateQteAnalObj(OBA,O,OQ,NQ,Ind) ;
END ;
{$ENDIF}

procedure ProrateQteAnalObj ( OBA : TObjAna ; O : TOBM ; OQ,NQ : Double ; Ind : integer ) ;
Var j,k  : integer ;
    QQ   : String3 ;
    AA   : TAA ;
    Diff,Coef,NQA,OQA,TotQA : Double ;
BEGIN
if OQ=0 then Exit ;
QQ:=O.GetMvt('E_QUALIFQTE'+IntToStr(Ind)) ; if QQ='' then Exit ;
Coef:=NQ/OQ ;
for j:=1 to MaxAxe do if OBA.AA[j]<>Nil then
    BEGIN
    AA:=OBA.AA[j] ;
    if Not CoherQualif(AA.L,QQ,Ind) then Continue ; TotQA:=0 ;
    for k:=0 to AA.L.Count-1 do
        BEGIN
        OQA:=GetMvtA(P_TV(AA.L[k]).F,'Y_QTE'+IntToStr(Ind)) ;
        NQA:=Arrondi(OQA*Coef,V_PGI.OkDecQ) ; PutMvtA(AA,P_TV(AA.L[k]).F,'Y_QTE'+IntToStr(Ind),NQA) ;
        TotQA:=TotQA+NQA ;
        END ;
    Diff:=Arrondi(NQ-TotQA,V_PGI.OkDecQ) ;
    if Diff<>0 then
       BEGIN
       k:=AA.L.Count-1 ; OQA:=GetMvtA(P_TV(AA.L[k]).F,'Y_QTE'+IntToStr(Ind)) ;
       NQA:=OQA+Diff ; PutMvtA(AA,P_TV(AA.L[k]).F,'Y_QTE'+IntToStr(Ind),NQA) ;
       END ;
    END ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/03/2004
Modifié le ... :   /  /
Description .. : - LG - 03/02/2004 - fct pour l'eAGL
Mots clefs ... :
*****************************************************************}
function CoherenceQualifTOB ( vTOB : TOB ; vStE_QUALIFQTE , vStY_QUALIFQTE : string ) : boolean ;
var
 i  : integer ;
begin
 result := false ;
 if ( vTOB = nil ) or ( vTOB.Detail = nil ) or ( vTOB.Detail.Count = 0 ) then exit ;
 for i := 0 to vTOB.Detail.Count - 1 do
  if vTOB.Detail[i].GetValue(vStY_QUALIFQTE) <> vStE_QUALIFQTE then exit ;
 result := true ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/03/2004
Modifié le ... :   /  /
Description .. : - LG - 03/02/2004 - fct pour l'eAGL
Mots clefs ... :
*****************************************************************}
procedure ProrateQteAnalTOB ( vTOB : TOB ;  vRdQte , vRdNQ : Double ; vStInd : string ) ;
var
 lRdCoef        : double ;
 i              : integer ;
 j              : integer ;
 k              : integer ;
 lTOBLigne      : TOB ;
 lRdQteAna      : Double ;
 lRdNQteAna     : double ;
 lRdTotalQteAna : double ;
 lStE_QUALIFQTE : string ;
 lRdDiff        : double ;
begin

 if EstSerie(S3) or EstSerie(S5) or ( vRdQte = 0 ) then exit ;
 if ( vTOB = nil ) or ( vTOB.Detail = nil ) or ( vTOB.Detail.Count = 0 ) then exit ;

 lStE_QUALIFQTE := vTOB.GetValue('E_QUALIFQTE' + vStInd ) ;
 if lStE_QUALIFQTE = '' then Exit ;

 lRdCoef := vRdNQ / vRDQte ;
 lRdTotalQteAna := 0 ;

 for i:=0 to vTOB.Detail.Count - 1 do
  begin
   lTOBLigne := vTOB.Detail[i] ;
   if not CoherenceQualifTOB(lTOBLigne,lStE_QUALIFQTE,'Y_QUALIFQTE' + vStInd ) then continue ;
   for j := 0 to lTOBLigne.Detail.Count - 1 do
    begin
     lRdQteAna := lTOBLigne.Detail[j].GetValue('Y_QTE' + vStInd ) ;
     lRdNQteAna := Arrondi( lRdQteAna * lRdCoef , V_PGI.OkDecV ) ;
     lTOBLigne.Detail[j].PutValue('Y_QTE' + vStInd , lRdNQteAna ) ;
     lRdTotalQteAna := lRdTotalQteAna + lRdNQteAna ;
    end ; // for j

   lRdDiff := Arrondi( vRdNQ  - lRdTotalQteAna , V_PGI.OkDecQ) ;
   if lRdDiff <> 0 then
    begin
     k := lTOBLigne.Detail.Count - 1 ;
     lTOBLigne.Detail[k].PutValue('Y_QTE' + vStInd ,  lTOBLigne.Detail[k].GetValue('Y_QTE' + vStInd ) + lRdDiff ) ;
    end ; // if

  end ; // for i

end ;



function CoherQualif ( L : TList ; QQ : String3 ; Ind : integer ) : boolean ;
Var k  : integer ;
    QA : String3 ;
BEGIN
CoherQualif:=False ; if L.Count<=0 then Exit ;
for k:=0 to L.Count-1 do
    BEGIN
    QA:=GetMvtA(P_TV(L[k]).F,'Y_QUALIFQTE'+IntToStr(Ind)) ;
    if QA<>QQ then Exit ;
    END ;
CoherQualif:=True ;
END ;

{function FabricReqComp ( fb : TFichierBase ) : String ;
Var SQL   : String ;
    i     : integer ;
BEGIN
Result:='' ; SQL:='' ;
Case fb of
   fbGene : BEGIN
            SQL:=' AND G_LETTRABLE=:G_L AND G_COLLECTIF=:G_C AND G_NATUREGENE=:G_N' ;
            for i:=1 to MaxAxe do SQL:=SQL+' AND G_VENTILABLE'+Inttostr(i)+'=:G_V'+Inttostr(i) ;
            END ;
   fbAux  : SQL:=' AND T_LETTRABLE=:T_L AND T_NATUREAUXI=:T_N' ;
   End ;
Result:=SQL ;
END ; }


procedure AttribParamsComp( Var FRM : TFRM ; OO : TObject ) ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    i    : integer ;
BEGIN
if OO=Nil then Exit ;
if OO is TGGeneral then
   BEGIN
   CGen:=TGGeneral(OO) ;
   if CGen.Lettrable then FRM.G_L:='X' else FRM.G_L:='-' ;
   if CGen.Collectif then FRM.G_C:='X' else FRM.G_C:='-' ;
   FRM.G_N:=CGen.NatureGene ;
   for i:=1 to 5 do
       BEGIN
       if CGen.Ventilable[i] then FRM.G_V[i]:='X' else FRM.G_V[i]:='-'  ;
       END ;
   END else if OO is TGTiers then
   BEGIN
   CAux:=TGTiers(OO) ;
   if CAux.Lettrable then FRM.T_L:='X' else FRM.T_L:='-' ;
   FRM.T_N:=CAux.NatureAux ;
   END ;
END ;


Function ExecReqINVNew ( fb : TFichierBase ; Var FRM : TFRM) : LongInt ;
Var
  SQL : String ;
begin
  Result := 0;

  if not EstTablePartagee( fbToTable( fb ) ) then
    begin
    // Gestion standard mode PCL OU si la table n'est pas partagée...
    Case fb of
       fbGene : SQL:='UPDATE GENERAUX SET G_TOTALDEBIT=G_TOTALDEBIT-'+StrFPoint(FRM.Deb)+', G_TOTALCREDIT=G_TOTALCREDIT-'+StrFPoint(FRM.Cre)+', '
                    +'G_TOTDEBE=G_TOTDEBE-'+StrFPoint(FRM.DE)+', G_TOTCREE=G_TOTCREE-'+StrFPoint(FRM.CE)+', '
                    +'G_TOTDEBS=G_TOTDEBS-'+StrFPoint(FRM.DS)+', G_TOTCRES=G_TOTCRES-'+StrFPoint(FRM.CS)+', '
                    +'G_TOTDEBP=G_TOTDEBP-'+StrFPoint(FRM.DP)+', G_TOTCREP=G_TOTCREP-'+StrFPoint(FRM.CP)+' WHERE G_GENERAL="'+FRM.Cpt+'"' ;
        fbAux : SQL:='UPDATE TIERS SET T_TOTALDEBIT=T_TOTALDEBIT-'+StrFPoint(FRM.Deb)+', T_TOTALCREDIT=T_TOTALCREDIT-'+StrFPoint(FRM.Cre)+', '
                    +'T_TOTDEBE=T_TOTDEBE-'+StrFPoint(FRM.DE)+', T_TOTCREE=T_TOTCREE-'+StrFPoint(FRM.CE)+', '
                    +'T_TOTDEBS=T_TOTDEBS-'+StrFPoint(FRM.DS)+', T_TOTCRES=T_TOTCRES-'+StrFPoint(FRM.CS)+', '
                    +'T_TOTDEBP=T_TOTDEBP-'+StrFPoint(FRM.DP)+', T_TOTCREP=T_TOTCREP-'+StrFPoint(FRM.CP)+' WHERE T_AUXILIAIRE="'+FRM.Cpt+'"' ;
       fbSect : SQL:='UPDATE SECTION SET S_TOTALDEBIT=S_TOTALDEBIT-'+StrFPoint(FRM.Deb)+', S_TOTALCREDIT=S_TOTALCREDIT-'+StrFPoint(FRM.Cre)+', '
                    +'S_TOTDEBE=S_TOTDEBE-'+StrFPoint(FRM.DE)+', S_TOTCREE=S_TOTCREE-'+StrFPoint(FRM.CE)+', '
                    +'S_TOTDEBS=S_TOTDEBS-'+StrFPoint(FRM.DS)+', S_TOTCRES=S_TOTCRES-'+StrFPoint(FRM.CS)+', '
                    +'S_TOTDEBP=S_TOTDEBP-'+StrFPoint(FRM.DP)+', S_TOTCREP=S_TOTCREP-'+StrFPoint(FRM.CP)+' WHERE S_AXE="'+FRM.Axe+'" AND S_SECTION="'+FRM.Cpt+'"' ;
        fbJal : SQL:='UPDATE JOURNAL SET J_TOTALDEBIT=J_TOTALDEBIT-'+StrFPoint(FRM.Deb)+', J_TOTALCREDIT=J_TOTALCREDIT-'+StrFPoint(FRM.Cre)+', '
                    +'J_TOTDEBE=J_TOTDEBE-'+StrFPoint(FRM.DE)+', J_TOTCREE=J_TOTCREE-'+StrFPoint(FRM.CE)+', '
                    +'J_TOTDEBS=J_TOTDEBS-'+StrFPoint(FRM.DS)+', J_TOTCRES=J_TOTCRES-'+StrFPoint(FRM.CS)+', '
                    +'J_TOTDEBP=J_TOTDEBP-'+StrFPoint(FRM.DP)+', J_TOTCREP=J_TOTCREP-'+StrFPoint(FRM.CP)+' WHERE J_JOURNAL="'+FRM.Cpt+'"' ;
       End ;
    Result := ExecuteSQL(SQL) ;
    end
  else
    // Gestion des cumuls MULTISOC
    begin
    if CUpdateCumulsMS ( fb, FRM.Cpt, FRM.Axe, -1*FRM.Deb, -1*FRM.Cre, -1*FRM.DP, -1*FRM.CP, -1*FRM.DE, -1*FRM.CE, -1*FRM.DS, -1*FRM.CS, False )
      then result := 1 ;
    end ;
END ;

function BudgetToIdent ( Q : TQuery ) : RMVT ;
Var M : RMVT ;
BEGIN
FillChar(M,Sizeof(M),#0) ;
M.Axe:=Q.FindField('BE_AXE').AsString ;
M.Etabl:=Q.FindField('BE_ETABLISSEMENT').AsString ;
M.Jal:=Q.FindField('BE_BUDJAL').AsString ;
M.CodeD:=V_PGI.DevisePivot ;
M.Simul:=Q.FindField('BE_QUALIFPIECE').AsString ;
M.Nature:=Q.FindField('BE_NATUREBUD').AsString ;
M.TypeSaisie:=Q.FindField('BE_TYPESAISIE').AsString ;
M.Num:=Q.FindField('BE_NUMEROPIECE').AsInteger ;
M.Valide:=(Q.FindField('BE_VALIDE').AsString='X') ;
Result:=M ;
END ;



function  SQLForIdent( fb : TFichierBase ) : string ;
var lStPref : string ;
begin
  Case fb of
     fbGene   : lStPref := 'E' ;
     fbSect   : lStPref := 'Y' ;
     fbBudgen : lStPref := 'BE' ;
     else       lStPref := 'E' ;
   end ;

  if fb=fbBudgen
    then result := 'BE_BUDJAL'
    else result := lStPref + '_JOURNAL' ;

  result := result + ', ' + lStPref + '_EXERCICE'
                   + ', ' + lStPref + '_QUALIFPIECE'
                   + ', ' + lStPref + '_DATECOMPTABLE'
                   + ', ' + lStPref + '_NUMEROPIECE'
                   + ', ' + lStPref + '_VALIDE'
                   + ', ' + lStPref + '_ETABLISSEMENT' ;

  if (fb<>fbBudgen) then
    result := result + ', ' + lStPref + '_DEVISE'
                     + ', ' + lStPref + '_NATUREPIECE'
                     + ', ' + lStPref + '_TAUXDEV'
                     + ', ' + lStPref + '_DATETAUXDEV' ;

  Case fb of
    fbSect : result := result + ', Y_AXE, Y_GENERAL, Y_NUMLIGNE' ;
    fbGene : result := result + ', E_NUMLIGNE, E_NUMECHE, E_GENERAL, E_ETATREVISION, E_MODESAISIE' ;
    end ;

end ;


function MvtToIdent( Q : TQuery ; fb : TFichierBase ; Totale : boolean ) : RMVT ;
Var M : RMVT ;
    Pf : String ;
    F  : TField ;
BEGIN
FillChar(M,Sizeof(M),#0) ;
Case fb of
   fbGene   : Pf:='E_' ;
   fbSect   : Pf:='Y_' ;
   fbBudgen   : Pf:='BE_' ;
   else Pf:='E_' ;
   END ;
if fb=fbBudgen then M.JAL:=Q.FindField('BE_BUDJAL').AsString else M.JAL:=Q.FindField(Pf+'JOURNAL').AsString ;
M.EXO:=Q.FindField(Pf+'EXERCICE').AsString ;
M.SIMUL:=Q.FindField(Pf+'QUALIFPIECE').AsString ;
M.DATEC:=Q.FindField(Pf+'DATECOMPTABLE').AsDateTime ;
M.NUM:=Q.FindField(Pf+'NUMEROPIECE').AsInteger    ;
M.VALIDE:=Q.FindField(Pf+'VALIDE').AsString='X'   ;
M.ETABL:=Q.FindField(Pf+'ETABLISSEMENT').AsString ;
if Pf='E_' then M.ModeSaisieJal:=Q.FindField('E_MODESAISIE').AsString ;
if (fb<>fbBudgen) then
   BEGIN
   M.CODED:=Q.FindField(Pf+'DEVISE').AsString        ;
   M.NATURE:=Q.FindField(Pf+'NATUREPIECE').AsString  ;
   M.TAUXD:=Q.FindField(Pf+'TAUXDEV').AsFloat ;
   M.DATETAUX:=Q.FindField(Pf+'DATETAUXDEV').AsDateTime ;
   END ;
Case fb of
   fbSect : BEGIN
            M.AXE:=Q.FindField('Y_AXE').AsString ; M.GENERAL:=Q.FindField('Y_GENERAL').AsString ;
            F:=Q.FindField('Y_NUMLIGNE') ; if F<>Nil then M.NumLigne:=F.AsInteger ;
            END ;
   fbGene : if Totale then
               BEGIN
               M.NumLigne:=Q.FindField('E_NUMLIGNE').AsInteger ; M.NumEche:=Q.FindField('E_NUMECHE').AsInteger ;
               M.General := Q.FindField('E_GENERAL').asString ;
               END ;
 END ;
//LG* 11/02/2002 correction : on ne passe dans cette fct que pour des rq sur la table ecriture
If (fb=fbGene) Then
  BEGIN
  If Q.FindField('E_ETATREVISION').asString='X' Then M.BloquePieceImport:=TRUE ;
  END ;
MVTTOIDENT:=M ;
END ;


function OBMToIdentManuel( O : TOBM ; Totale : boolean ; NatPiece : String3) : RMVT ;
Var M : RMVT ;
    S1 : String3 ;
BEGIN
FillChar(M,Sizeof(M),#0) ;
If NatPiece='' Then S1:=O.GetMvt('E_NATUREPIECE') Else S1:=NatPiece ;
M.Etabl:=O.GetMvt('E_ETABLISSEMENT') ;
M.Jal:=O.GetMvt('E_JOURNAL')         ; M.Exo:=O.GetMvt('E_EXERCICE') ;
M.DateC:=O.GetMvt('E_DATECOMPTABLE') ; M.Nature:=S1 ;
M.Num:=O.GetMvt('E_NUMEROPIECE')     ; M.Simul:=O.GetMvt('E_QUALIFPIECE') ;
M.CodeD:=O.GetMvt('E_DEVISE')        ; M.DateTaux:=O.GetMvt('E_DATETAUXDEV') ;
M.TauxD:=O.GetMvt('E_TAUXDEV')       ; M.Valide:=(O.GetMvt('E_VALIDE')='X') ;
if Totale then
   BEGIN
   M.NumLigne:=O.GetMvt('E_NUMLIGNE') ; M.NumEche:=O.GetMvt('E_NUMECHE') ;
   END ;
OBMToIdentManuel:=M ;
END ;

Function CompleteReseau ( ts : TTypeSais ; O : TOBM ) : String ;
BEGIN
Result:='' ; if O=Nil then Exit ;
Case ts of
   tsGene : Result:=' AND E_REFPOINTAGE="'+O.GetMvt('E_REFPOINTAGE')+'" AND E_NIVEAURELANCE='+Inttostr(O.GetMvt('E_NIVEAURELANCE'))+' ' ;
   END ;
END ;

{$IFNDEF EAGLSERVER}
procedure ComCom1 ( G : THGrid ; O : TOB ) ;
BEGIN
ComComLigne(G,O,G.RowCount-1) ;
END ;

procedure ComComLigne ( G : THGrid ; O : TOB ; Lig : integer ) ;
Var StC,STitreCol,StCell : String ;
    C  : integer ;
    V  : Variant ;
BEGIN
{ FQ 21458 BVE 19.09.07 }
if G.Titres.Count > 0 then
   StC := G.Titres[0]
else
   Exit;
{ END FQ 21458 }
C:=0 ;
Repeat
 STitreCol:=ReadTokenSt(StC) ; if STitreCol='' then Break ;
 if STitreCol[1]='$' then System.Delete(STitreCol,1,1) ;
 V:=O.GetValue(STitreCol) ;
 if ((G.ColFormats[C]<>'') and (VarType(V)=VarDouble)) then StCell:=FormatFloat(G.ColFormats[C],VarAsType(V,VarDouble))
                                                       else StCell:=VarAsType(V,VarString) ;
 G.Cells[C,Lig]:=StCell ;
 inc(C) ;
Until ((StC='') or (STitreCol='') or (C>=G.ColCount)) ;
END ;
{$ENDIF}

function Scenario2Comp ( Scenario : TOB ) : String ;
Var StComp : String ;
BEGIN
StComp:='----------' ;
if Scenario.GetValue('SC_REFINTERNE')='X' then StComp[1]:='X' ;
if Scenario.GetValue('SC_LIBELLE')='X' then StComp[2]:='X' ;
if Scenario.GetValue('SC_REFEXTERNE')='X' then StComp[3]:='X' ;
if Scenario.GetValue('SC_DATEREFEXTERNE')='X' then StComp[4]:='X' ;
if Scenario.GetValue('SC_AFFAIRE')='X' then StComp[5]:='X' ;
if Scenario.GetValue('SC_REFLIBRE')='X' then StComp[6]:='X' ;
Result:=StComp ;
END ;

Function ConvertJoker(nom : string): String  ;
BEGIN
Result:='' ;
while Pos('?', Nom) > 0 do Nom[Pos('?', Nom)] := '_' ;
while Pos('*', Nom) > 0 do Nom[Pos('*', Nom)] := '%' ;
Result:=Nom ;
END ;

{$IFNDEF NOVH}
{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
procedure RenseigneHTByTVA ( CGen : TGGeneral ; RegTVA,CodeTva,CodeTPF : String3 ; SoumisTpf,Achat,EncON : boolean ; Var CoherTva : boolean ) ;
Var StC : String ;
BEGIN
StC:='' ;
if EncON then StC:=Tva2Encais(RegTVA,CodeTva,Achat) else StC:=Tva2Cpte(RegTVA,CodeTva,Achat) ;
if Not Presence('GENERAUX','G_GENERAL',StC) then StC:='' ;
CGen.CpteTVA:=StC ; CGen.TauxTVA:=Tva2Taux(RegTVA,CodeTva,Achat) ;
if SoumisTpf then
   BEGIN
   StC:=Tpf2Cpte(RegTVA,CodeTpf,Achat) ;
   if Not Presence('GENERAUX','G_GENERAL',StC) then StC:='' ;
   CGen.CpteTPF:=StC ;
   CGen.TauxTPF:=Tpf2Taux(RegTVA,CodeTpf,Achat) ;
   END ;
END ;
{$ENDIF}
{$ENDIF}
{$ENDIF}

function IncrementeNumeroBudget : longint ;
Var Facturier : String3 ;
    Q         : TQuery ;
    NumPieceInt : longint ;
BEGIN
{Result:=0 ;} NumPieceInt:=0;
Q:=OpenSQL('Select SH_SOUCHE from SOUCHE Where SH_TYPE="BUD"',True,-1,'',true) ;
if Not Q.EOF then BEGIN
  Facturier:=Q.Fields[0].AsString ;
  SetIncNum(EcrBud,Facturier,NumPieceInt,0) ;
END ;
Ferme(Q) ;
Result:=NumPieceInt ;
END ;

procedure ECCANA ( T : TDataSet ; CodeD : String3 ) ;
BEGIN
{$IFDEF EAGLCLIENT}
T.PutValue('Y_DEBITDEV', 0);
T.PutValue('Y_CREDITDEV', 0);
T.PutValue('Y_DEVISE', CodeD);
{$ELSE}
T.FindField('Y_DEBITDEV').AsFloat:=0 ;
T.FindField('Y_CREDITDEV').AsFloat:=0 ;
T.FindField('Y_DEVISE').AsString:=CodeD ;
{$ENDIF}
END ;

procedure ECCECR ( T : TDataSet ; CodeD : String3 ) ;
BEGIN
{$IFDEF EAGLCLIENT}
T.PutValue('E_DEBITDEV', 0);
T.PutValue('E_CREDITDEV', 0);
T.PutValue('E_DEVISE', CodeD);
{$ELSE}
T.FindField('E_DEBITDEV').AsFloat:=0 ;
T.FindField('E_CREDITDEV').AsFloat:=0 ;
T.FindField('E_DEVISE').AsString:=CodeD ;
{$ENDIF}
END ;

procedure RemplirOFromM ( M : RMVT ; ListeSel : TList ) ;
Var QEcr : TQuery ;
    O    : TOBM ;
BEGIN
QEcr:=OpenSQL('Select * from Ecriture Where '+WhereEcriture(tsGene,M,False)+' AND E_NUMLIGNE='+InttoStr(M.NumLigne)+' order by E_NUMECHE',True,-1,'',true) ;
While Not QEcr.EOF do
   BEGIN
   O:=TOBM.Create(EcrGen,'',False) ; O.ChargeMvt(QEcr) ;
   ListeSel.Add(O) ;
   QEcr.Next ;
   END ;
Ferme(QEcr) ;
END ;


function EstFormule ( St : String ) : boolean ;
BEGIN
Result:=Pos('{',St)>0 ;
if Not Result then Result:=Pos('}',St)>0 ;
if Not Result then Result:=Pos('[',St)>0 ;
if Not Result then Result:=Pos(']',St)>0 ;
END ;

Function ModifRIBOBM (O : TOB ; Maj,Treso : boolean ; Aux : String ; IsAux : Boolean; bCFONBIntern : Boolean = False) : boolean ;
{$IFNDEF EAGLSERVER}
  Var
    Auxi   : String ;
    {$IFDEF COMPTA}
      Nb : integer;
      M  : RMVT;
      SQL: string;
    {$ENDIF COMPTA}
    Num : integer ;
    Q      : TQuery ;
    RIB,Etab,Guichet,NumCompte,Cle,Dom,Iban,Pays : String ;
{$ENDIF !EAGLSERVER}
BEGIN
Result:=False ;
{$IFNDEF EAGLSERVER}
If IsAux Then Auxi:=O.GetString('E_AUXILIAIRE') Else Auxi:=O.GetString('E_GENERAL') ;
if (Auxi='') And (Not Treso) then Exit ;
If Auxi='' Then Auxi:=Aux ; If Auxi='' Then Exit ;
RIB:=O.GetString('E_RIB') ;
if VH^.PaysLocalisation<>CodeISOES then
   Begin
   Num:=FicheRIB_AGL(Auxi,taModif,True,RIB,IsAux) ; if Num<=0 then Exit ;
   Q:=OpenSQL('Select * from RIB Where R_AUXILIAIRE="'+Auxi+'" AND R_NUMERORIB='+IntToStr(Num),True,-1,'',true) ;
   if Not Q.EOF then
      BEGIN
      Etab:=Q.FindField('R_ETABBQ').AsString ;
      Guichet:=Q.FindField('R_GUICHET').AsString ;
      NumCompte:=Q.FindField('R_NUMEROCOMPTE').AsString ;
      Cle:=Q.FindField('R_CLERIB').AsString ;
      Dom:=Q.FindField('R_DOMICILIATION').AsString ;
      Iban:=Q.FindField('R_CODEIBAN').AsString ;
      Pays:=Q.FindField('R_PAYS').AsString ;
      END ;
   Ferme(Q) ;
   if ((codeisodupays(Pays)<>'FR') or (bCFONBIntern)) then // modif CAY 11/07/2003
     RIB:='*'+Iban
   else
     RIB:=EncodeRIB(Etab,Guichet,NumCompte,Cle,Dom) ;
   End ;
{$IFDEF COMPTA}
if (VH^.PaysLocalisation<>CodeISOES) or (ModifLeRib(RIB,Auxi)) then
  Begin
  O.PutValue('E_RIB',RIB) ;
  if Maj then
     BEGIN
     M:=TobToIdent(O,True) ;
     {JP 16/11/07 : FQ 21847 : Gestion de E_UTILISATEUR}
     SQL:='UPDATE ECRITURE SET E_RIB = "' + RIB + '", E_UTILISATEUR = "' +
           V_PGI.User + '", E_DATEMODIF = "' + UsTime(NowH) + '" WHERE ' + WhereEcriture(tsGene, M, True);
     Nb:=ExecuteSQL(SQL) ; if Nb=1 then Result:=True ;
     END ;
   End ;
{$ENDIF}
{$ENDIF}
END ;
{$ENDIF}
{$ENDIF}

Function EstCollFact ( Coll : String ) : boolean ;
Var St,StC : String ;
BEGIN
{#TVAENC}
Result:=False ;
if Not GetParamsoc('SO_OUITVAENC') then Exit ;
St := GetParamSoc('SO_COLLCLIENC');
if St<>'' then
   BEGIN
   if St[Length(St)]<>';' then St:=St+';' ;
   Repeat
    StC:=ReadTokenSt(St) ; if StC='' then Break ;
    if Copy(Coll,1,Length(StC))=StC then BEGIN Result:=True ; Break ; END ;
   Until ((St='') or (StC='')) ;
   END ;
if Not Result then
   BEGIN
   St := GetParamsoc('SO_COLLFOUENC');
   if St<>'' then
      BEGIN
      if St[Length(St)]<>';' then St:=St+';' ;
      Repeat
       StC:=ReadTokenSt(St) ; if StC='' then Break ;
       if Copy(Coll,1,Length(StC))=StC then BEGIN Result:=True ; Break ; END ;
      Until ((St='') or (StC='')) ;
      END ;
   END ;
END ;

{$IFNDEF NOVH}

Function EstJalFact ( Jal : String3 ) : boolean ;
BEGIN
Result:=(Pos(';'+Jal+';',VH^.ListeNatJal)>0) ;
END ;
{$ENDIF}

{$IFNDEF EAGLSERVER}
procedure ColorOpposeEuro ( G : THGrid ; ModeOppose,EnDevise : boolean ) ;
Var Vis : boolean ;
    FF  : TForm ;
    II  : TImage ;
    St  : String ;
BEGIN
Vis:=False ; St:=G.Cells[0,0] ;
if ModeOppose then
   BEGIN
   G.FixedColor:=cl3DLight ; G.TitleBold:=True  ;
   if Not GetParamsoc('SO_TENUEEURO') then Vis:=True ;
   END else
   BEGIN
     {$IFDEF GCGC}  //JSI conflit Look V6
     if G.FixedColor=cl3DLight then G.FixedColor:=clBtnFace ;
     {$ELSE}
     G.FixedColor:=clBtnFace ;
    {$ENDIF}
     G.TitleBold:=False ;
   if ((GetParamsoc('SO_TENUEEURO')) and (Not EnDevise)) then Vis:=True ;
   END ;
FF:=TForm(G.Owner) ; II:=TImage(FF.FindComponent('ISIGNEEURO')) ;
if II<>Nil then II.Visible:=Vis ;
G.Cells[0,0]:=St ;
END ;

Procedure DelTabsSerie ( TA : TTabControl ) ;
BEGIN
// JLD5Axes
if EstSerie(S3) then
   BEGIN
   TA.Tabs.Delete(4) ; TA.Tabs.Delete(3) ; TA.Tabs.Delete(2) ;
   TA.Tabs.Delete(1) ;
   END ;
if EstSerie(S3) then TA.Visible:=False ;
END ;
{$ENDIF}

function TOBToIdent ( O : TOB ; Totale : boolean ) : RMVT ;
Var M : RMVT ;
BEGIN //YMO 02/01/2007 Suppression des GetValue
FillChar(M,Sizeof(M),#0) ;
M.Etabl:=O.GetString('E_ETABLISSEMENT') ;
M.Jal:=O.GetString('E_JOURNAL')         ;
M.Exo:=O.GetString('E_EXERCICE') ;
M.DateC:=O.GetDateTime('E_DATECOMPTABLE') ;
M.Nature:=O.GetString('E_NATUREPIECE') ;
M.Num:=O.GetInteger('E_NUMEROPIECE')     ;
M.Simul:=O.GetString('E_QUALIFPIECE') ;
M.CodeD:=O.GetString('E_DEVISE');
M.DateTaux:=O.GetDateTime('E_DATETAUXDEV') ;
M.TauxD:=O.GetDouble('E_TAUXDEV')       ;
M.Valide:=(O.GetString('E_VALIDE')='X') ;
M.ModeSaisieJal:=O.GetString('E_MODESAISIE') ;
if Totale then
   BEGIN
   M.NumLigne:=O.GetInteger('E_NUMLIGNE') ; M.NumEche:=O.GetInteger('E_NUMECHE') ;
   M.General := O.GetString('E_GENERAL') ;
   END ;
Result:=M ;
END ;

Function WhereEcriture ( ts : TTypeSais ; M : RMVT ; Totale : boolean ; ChargeUnBordereau : Boolean = FALSE) : String ;
BEGIN
Case ts of
   tsGene,tsTreso,tsPointage,tsLettrage :
              BEGIN
              (* GP le 30/05/2002
              Result:='E_JOURNAL="'+M.JAL+'" AND E_EXERCICE="'+M.EXO+'"'
              +' AND E_DATECOMPTABLE="'+UsDateTime(M.DateC)+'" AND E_NUMEROPIECE='+InttoStr(M.Num)
              +' AND E_QUALIFPIECE="'+M.Simul+'" AND E_NATUREPIECE="'+M.Nature+'"' ;
              if Totale then Result:=Result+' AND E_NUMLIGNE='+IntToStr(M.NumLigne)+' AND E_NUMECHE='+IntToStr(M.NumEche) ;
              *)
              If ChargeUnBordereau Then
                BEGIN
                Result:='E_JOURNAL="'+M.JAL+'" AND E_EXERCICE="'+M.EXO+'"'
                +' AND E_PERIODE='+IntToStr(GetPeriode(M.DateC))+' AND E_NUMEROPIECE='+InttoStr(M.Num)
                +' AND E_QUALIFPIECE="'+M.Simul+'" ' ;
                END Else
                BEGIN
                Result:='E_JOURNAL="'+M.JAL+'" AND E_EXERCICE="'+M.EXO+'"'
                +' AND E_DATECOMPTABLE="'+UsDateTime(M.DateC)+'" AND E_NUMEROPIECE='+InttoStr(M.Num)
                +' AND E_QUALIFPIECE="'+M.Simul+'" AND E_NATUREPIECE="'+M.Nature+'"' ;
                END ;
              if Totale then Result:=Result+' AND E_NUMLIGNE='+IntToStr(M.NumLigne)+' AND E_NUMECHE='+IntToStr(M.NumEche) ;
              END ;
   tsBudget : BEGIN
              Result:='BE_BUDJAL="'+M.Jal+'" AND BE_NATUREBUD="'+M.Nature+'" AND BE_QUALIFPIECE="'+M.Simul+'"'
              +' AND BE_NUMEROPIECE='+IntToStr(M.Num) ;
              if Totale then Result:=Result+' AND BE_BUDGENE="'+M.General+'" AND BE_BUDSECT="'+M.Section+'" AND BE_AXE="'+M.Axe+'"' ;
              END ;
         else BEGIN
              {tsAnal, tsODA}
              Result:='Y_JOURNAL="'+M.JAL+'" AND Y_EXERCICE="'+M.EXO+'"'
              +' AND Y_DATECOMPTABLE="'+UsDateTime(M.DateC)+'" AND Y_NUMEROPIECE='+InttoStr(M.Num)
              +' AND Y_QUALIFPIECE="'+M.Simul+'" AND Y_NATUREPIECE="'+M.Nature+'"' ;
              if ts=tsODA then
                 BEGIN
                 Result:=Result+' AND Y_AXE="'+M.AXE+'" AND Y_GENERAL="'+M.GENERAL+'"' ;
                 if M.NumLigne>0 then Result:=Result+' AND Y_NUMLIGNE='+IntToStr(M.NumLigne) ;
                 END ;
              END ;
   END ;
END ;

{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
Procedure InitTableTemporaire(TypeTraitement : String) ;
BEGIN
{$IFDEF CCMP}
If TypeTraitement='' Then ExecuteSQL('DELETE FROM CPMPTEMPOR WHERE CTT_USER="'+V_PGI.USER+'" ')
                     Else ExecuteSQL('DELETE FROM CPMPTEMPOR WHERE CTT_USER="'+V_PGI.USER+'" AND CTT_TYPETRAITEMENT="'+TypeTraitement+'" ') ;
{$ENDIF}
END ;
{$ENDIF}
{$ENDIF}

Function SmpToString(smp : TSuiviMP) : String ;
BEGIN
Result:='' ;
Case smp of
   smpEncTraEnc   : Result:='CTE' ;
   smpEncTraEsc   : Result:='CTS' ;
//     smpDecChqEdt : Result:='FCH' ;
   smpDecVirEdt   : Result:='FVI' ;
   smpDecVirInEdt : Result:='FVN' ;
   smpEncPreEdt   : Result:='CPR' ;
   smpDecBorDec,smpDecBorEsc : Result:='FBO' ;
   END ;
END ;

{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
Function AlimTableTemporaire(O : TOBM ; LeTypeTraitement : String ; smp : tSuiviMP ; ZL1,ZL2 : String) : Boolean ;
Var Q : TOB ;
    St : String ;
    TypeTraitement : String ;
BEGIN
  Result:=FALSE ;
{$IFNDEF CCMP}
Exit ;
{$ENDIF}
  If EstSerie(S3) Then Exit ;
  TypeTraitement := SmpToString(smp) ;
  If TypeTraitement='' Then
    TypeTraitement:=LeTypeTraitement ;
  If TypeTraitement='' Then Exit ;
  St:='SELECT * FROM CPMPTEMPOR'
             + ' WHERE CTT_USER="' + V_PGI.User + '"'                   // Clé
               + ' AND CTT_TYPETRAITEMENT="' + TypeTraitement + '"'     // Clé
//               + ' AND CTT_EXERCICE="' + O.GetMvt('E_EXERCICE') + '"'
//               + ' AND CTT_JOURNAL="'+O.GetMvt('E_JOURNAL')+'"'
               + ' AND CTT_DATECOMPTABLE="'+USDateTime(O.GetMvt('E_DATECOMPTABLE'))+'"' // Clé
               + ' AND CTT_NUMEROPIECE='+IntToStr(O.GetMvt('E_NUMEROPIECE')) // Clé
               + ' AND CTT_NUMLIGNE='+IntToStr(O.GetMvt('E_NUMLIGNE'))       // Clé
               + ' AND CTT_NUMECHE='+IntToStr(O.GetMvt('E_NUMECHE')) ;        // Clé
//               + ' AND CTT_QUALIFPIECE="N" ' ;

//Q:=OpenSQL('SELECT * FROM CPMPTEMPOR WHERE CTT_USER="'+W_W+'" ',FALSE) ; // ?? requete d'origine ??
  If not ExisteSQL(St) Then
    BEGIN
    Q := TOB.Create('CPMPTEMPOR', nil, -1) ;
    Q.InitValeurs ;
    Q.PutValue('CTT_USER',            V_PGI.User) ;
    Q.PutValue('CTT_TYPETRAITEMENT',  TypeTraitement ) ;
    Q.PutValue('CTT_DATECOMPTABLE',   O.GetMvt('E_DATECOMPTABLE') ) ;
    Q.PutValue('CTT_NUMEROPIECE',     O.GetMvt('E_NUMEROPIECE') ) ;
    Q.PutValue('CTT_NUMLIGNE',        O.GetMvt('E_NUMLIGNE') ) ;
    Q.PutValue('CTT_NUMECHE',         O.GetMvt('E_NUMECHE') ) ;
    Q.PutValue('CTT_EXERCICE',        O.GetMvt('E_EXERCICE') ) ;
    Q.PutValue('CTT_JOURNAL',         O.GetMvt('E_JOURNAL') ) ;
    Q.PutValue('CTT_QUALIFPIECE',     O.GetMvt('E_QUALIFPIECE') ) ;
    Q.PutValue('CTT_GENERAL',         O.GetMvt('E_GENERAL') ) ;
    Q.PutValue('CTT_AUXILIAIRE',      O.GetMvt('E_AUXILIAIRE') ) ;
    Q.PutValue('CTT_ZONELIBRE1',      ZL1 ) ;
    Q.PutValue('CTT_ZONELIBRE2',      ZL2 ) ;
    Q.InsertDB(nil) ;
    Result:=TRUE ;
    END ;

END ;
{$ENDIF}
{$ENDIF}

{$IFDEF ANPRO}
// Gendreau laurent compta 25/02/2004 pas utilise de la compta a mettre ailleur
{
Function  TraiteANPRO(DatePiece : tDateTime) : Boolean ;
BEGIN
Result:=FALSE ;
If VH^.EnCours.EtatCpta<>'CPR' Then Exit ;
If (DatePiece>VH^.EnCours.Fin) Or (DatePiece<VH^.EnCours.Deb) Then Exit ;
Result:=TRUE ;
END ;}
{$ENDIF}

function OBMToIdent( O : TOBM ; Totale : boolean ) : RMVT ;
Var M : RMVT ;
BEGIN
FillChar(M,Sizeof(M),#0) ;
if O.Ident=EcrGen then
   BEGIN
   M.Etabl:=O.GetMvt('E_ETABLISSEMENT') ;
   M.Jal:=O.GetMvt('E_JOURNAL')         ; M.Exo:=O.GetMvt('E_EXERCICE') ;
   M.DateC:=O.GetMvt('E_DATECOMPTABLE') ; M.Nature:=O.GetMvt('E_NATUREPIECE') ;
   M.Num:=O.GetMvt('E_NUMEROPIECE')     ; M.Simul:=O.GetMvt('E_QUALIFPIECE') ;
   M.CodeD:=O.GetMvt('E_DEVISE')        ; M.DateTaux:=O.GetMvt('E_DATETAUXDEV') ;
   M.TauxD:=O.GetMvt('E_TAUXDEV')       ; M.Valide:=(O.GetMvt('E_VALIDE')='X') ;
   M.ModeSaisieJal:=O.GetMvt('E_MODESAISIE') ;
   if Totale then
      BEGIN
      M.NumLigne:=O.GetMvt('E_NUMLIGNE') ; M.NumEche:=O.GetMvt('E_NUMECHE') ;
      END ;
   END else if O.Ident=EcrBud then
   BEGIN
   M.Jal:=O.GetMvt('BE_BUDJAL')       ; M.Axe:=O.GetMvt('BE_AXE') ;
   M.Nature:=O.GetMvt('BE_NATUREBUD') ; M.Etabl:=O.GetMvt('BE_ETABLISSEMENT') ;
   M.Num:=O.GetMvt('BE_NUMEROPIECE')  ; M.Simul:=O.GetMvt('BE_QUALIFPIECE') ;
   M.Valide:=(O.GetMvt('BE_VALIDE')='X') ;
   M.CodeD:=V_PGI.DevisePivot ; M.TauxD:=1.0 ;
   if Totale then
      BEGIN
      M.Exo:=O.GetMvt('BE_EXERCICE') ;
      M.DateC:=O.GetMvt('BE_DATECOMPTABLE') ;
      M.General:=O.GetMvt('BE_BUDGENE') ;
      M.Section:=O.GetMvt('BE_BUDSECT') ;
      M.DateTaux:=M.DateC ;
      END ;
   END else Exit ;
OBMToIdent:=M ;
END ;

procedure _AjouteAnoCompte ( TS : TList ; vTOBEcr : TOB ; Inverse : boolean ; vBoTous : boolean = false ) ;
var
 i             : integer ;
 lpT           : ^TFRM ;
 lRdSolde      : double ;
 lRdSoldeDev   : double ;
 lRdDebit      : double ;
 lRdCredit     : double ;
 lRdDebitDev   : double ;
 lRdCreditDev  : double ;
begin


 if TS = nil then exit ;

 lRdDebit     := vTOBEcr.GetValue('E_DEBIT') ;
 lRdCredit    := vTOBEcr.GetValue('E_CREDIT') ;
 lRdDebitDev  := vTOBEcr.GetValue('E_DEBITDEV') ;
 lRdCreditDev := vTOBEcr.GetValue('E_CREDITDEV') ;


 if Inverse then
  begin
   lRdDebit     := lRdDebit * (-1) ;
   lRdDebitDev  := lRdDebitDev * (-1) ;
   lRdCredit    := lRdCredit * (-1) ;
   lRdCreditDev := lRdCreditDev * (-1) ;
  end ;

 for i := 0 to TS.Count - 1 do
  begin

   lpT := TS.Items[i] ;

   if ( lpT^.Cpt      = vTOBEcr.GetValue('E_GENERAL') )       and
      ( lpT^.Axe      = vTOBEcr.GetValue('E_AUXILIAIRE') ) and
      ( lpT^.Etabliss = vTOBEcr.GetValue('E_ETABLISSEMENT') ) and
      ( lpT^.Dev      = vTOBEcr.GetValue('E_DEVISE') ) then
    begin
     lRdSolde    := ( lpT^.Deb + lRdDebit ) - ( lpT^.Cre + lRdCredit ) ;
     lRdSoldeDev := ( lpT^.Deb + lRdDebitDev ) - ( lpT^.Cre + lRdCreditDev ) ;

     if Arrondi(lRdSolde,4) > 0 then
      begin
       lpT^.Deb    := lRdSolde ;
       lpT^.DebDev := lRdSoldeDev ;
       lpT^.DS     := lRdSolde ;
       lpT^.Cre    := 0 ;
       lpT^.CS     := 0 ;
       lpT^.CreDev := 0 ;
      end
       else
        begin
         lpT^.Deb    := 0 ;
         lpT^.DS     := 0 ;
         lpT^.DebDev := 0 ;
         lpT^.Cre    := lRdSolde*(-1) ;
         lpT^.CreDev := lRdSoldeDev*(-1) ;
         lpT^.CS     := lRdSolde*(-1) ;
        end ;
     exit ;
    end ; // if

  end ; // for

 New(lpT);
 FillChar(lpT^,sizeOf(lpT^),#0);
 if vBoTous then
  lpT^.fb := fbAxe3
  else
   lpT^.fb         := fbAxe2 ;
 lpT^.Cpt        := vTOBEcr.GetValue('E_GENERAL') ;
 if length(trim(vTOBEcr.GetValue('E_AUXILIAIRE'))) > 0 then
  lpT^.Axe := vTOBEcr.GetValue('E_AUXILIAIRE')
   else
    lpT^.Axe := '' ;
 lpT^.Deb        := lRdDebit ;
 lpT^.Cre        := lRdCredit ;
 lpT^.DebDev     := lRdDebitDev ;
 lpT^.CreDev     := lRdCreditDev ;
 lpT^.CS         := lRdCredit ;
 lpT^.DS         := lRdDebit ;
 lpT^.Etabliss   := vTOBEcr.GetValue('E_ETABLISSEMENT') ;
 lpT^.Dev        := vTOBEcr.GetValue('E_DEVISE') ;
 lpT^.Exo        := vTobEcr.GetString('E_EXERCICE');
 // GCO - 04/07/2006
 lpT^.DateD      := vTobEcr.GetDateTime('E_DATECOMPTABLE');

 TS.Add(lpT);

end ;

procedure _AjouteAnoChPr ( TS : TList ; vTOBEcr : TOB ; Inverse : boolean ) ;
var
 i                : integer ;
 lpT              : ^TFRM ;
 lRdSolde         : double ;
 lRdDebit         : double ;
 lRdCredit        : double ;
 lRdDebitDev      : double ;
 lRdCreditDev     : double ;
 lRdSoldeDev      : double ;
begin

 lRdDebit     := vTOBEcr.GetValue('E_DEBIT') ;
 lRdCredit    := vTOBEcr.GetValue('E_CREDIT') ;
 lRdDebitDev  := vTOBEcr.GetValue('E_DEBITDEV') ;
 lRdCreditDev := vTOBEcr.GetValue('E_CREDITDEV') ;

 if Inverse then
  begin
   lRdDebit     := lRdDebit * (-1) ;
   lRdCredit    := lRdCredit * (-1) ;
   lRdDebitDev  := lRdDebitDev * (-1) ;
   lRdCreditDev := lRdCreditDev * (-1) ;
  end ;

 for i := 0 to TS.Count - 1 do
  begin

   lpT := TS.Items[i] ;

   if ( lpT^.fb = fbAxe1 ) then
    begin
     lRdSolde    := ( lpT^.Cre + lRdCredit ) - ( lpT^.Deb + lRdDebit ) ;
     lRdSoldeDev := ( lpT^.CreDev + lRdCreditDev ) - ( lpT^.DebDev + lRdDebitDev ) ;
     if Arrondi(lRdSolde,4) > 0 then
      begin
       lpT^.Deb    := 0 ;
       lpT^.Cre    := lRdSolde ;
       lpT^.DebDev := 0 ;
       lpT^.CreDev := lRdSoldeDev ;
       lpT^.DS     := 0 ;
       lpT^.CS     := lRdSolde ;
      end
       else
        begin
         lpT^.Deb    := lRdSolde * (-1) ;
         lpT^.Cre    := 0 ;
         lpT^.DebDev := lRdSoldeDev * (-1) ;
         lpT^.CreDev := 0 ;
         lpT^.DS     := lRdSolde * (-1) ;
         lpT^.CS     := 0 ;
        end ;
     exit ;
    end ; // if

  end ;

  New(lpT);
  FillChar(lpT^,sizeOf(lpT^),#0);
  lpT^.fb         := fbAxe1 ;
  lpT^.Deb        := lRdDebit ;
  lpT^.Cre        := lRdCredit ;
  lpT^.DebDev     := lRdDebitDev ;
  lpT^.CreDev     := lRdCreditDev ;
  lpT^.DS         := lRdDebit ;
  lpT^.CS         := lRdCredit ;
  lpT^.Etabliss   := vTOBEcr.GetValue('E_ETABLISSEMENT') ;
  lpT^.Dev        := vTOBEcr.GetValue('E_DEVISE') ;
  // GCO - 04/07/2006
  lpT^.DateD      := vTobEcr.GetDateTime('E_DATECOMPTABLE');
  lpT^.Exo        := vTobEcr.GetString('E_EXERCICE');
  TS.Add(lpT);

end ;



procedure AjouteAno ( TS : TList ; vTOBEcr : TOB ; NatGene : string ; Inverse : boolean ) ;
var
 lBoBDSDyna : Boolean ;
 lBoAnoDyna : boolean ;
begin

 if TS = nil then exit ;

 lBoBDSDyna := GetParamSocSecur('SO_CPBDSDYNA', False);
 lBoAnoDyna := _GetParamSocSecur('SO_CPANODYNA',false) and ( ctxExercice.Suivant.Code <> '' ) and ( vTOBEcr.GetValue('E_EXERCICE') = ctxExercice.EnCours.Code ) ;

 if not ( lBoAnoDyna or lBoBDSDyna ) then Exit ;

 if ( NatGene = 'CHA' ) or ( NatGene = 'PRO' ) then
 begin

   _AjouteAnoChPr(TS,vTOBEcr,Inverse);

  if lBoBDSDyna then
    _AjouteAnoCompte(TS,vTOBEcr,Inverse,true);
 end
 else
   _AjouteAnoCompte(TS,vTOBEcr,Inverse) ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 18/05/2004
Modifié le ... :   /  /
Description .. : - LG - 18/05/2004 - modif pour l'acceleration des comptes
Suite ........ : en saisie, ajout de cptecegid
Mots clefs ... :
*****************************************************************}
procedure Ajoute ( TS : TList ; St : String3 ; Cpte : String17 ; D,C : double ; NumP,NumL : Integer ; DateP : tDateTime ; CptCegid : string = '' ; YtcTiers : string = '' ) ;
var
 i : integer ;
 T : T_SC ;

 procedure _Recopie ;
  begin
   T.Identi:=St ; T.Cpte:=Cpte ;
   T.Debit:=T.Debit+D ; T.Credit:=T.Credit+C ;
   T.NumP:=NumP ; T.NumL:=NumL ; T.DateP:=DateP ;
   if CptCegid <> '' then
    begin
     T.CptCegid:=CptCegid ;
     T.YtcTiers:=YtcTiers ;
   end ;
  end ;

begin

 if Trim(Cpte) = '' then Exit ;

 for i := 0 to TS.Count-1 do
  begin
    T := T_SC(TS[i]) ;
    if ((T.Identi=St) and (T.Cpte=Cpte)) then
     begin
      _Recopie ;
      exit ;
     end ;
  end ; // for

 T := T_SC.Create ;
 _Recopie ;
 TS.Add(T) ;

end ;

procedure VideTFRM ( var Value : TFRM ) ;
begin
 Value.Cpt         := #0 ;
 Value.Axe         := #0 ;
 Value.Deb         := 0 ;
 Value.Cre         := 0 ;
 Value.DE          := 0 ;
 Value.CE          := 0 ;
 Value.DS          := 0 ;
 Value.CS          := 0 ;
 Value.DP          := 0 ;
 Value.CP          := 0 ;
 Value.DebDev      := 0 ;
 Value.CreDev      := 0 ;
 Value.CreDev      := 0 ;
 Value.DateD       := iDate1900 ;
 Value.NumD        := 0 ;
 Value.LigD        := 0 ;
 Value.NumClo      := 0 ;
 Value.NumAno      := 0 ;
 Value.G_L         := #0 ;
 Value.G_C         := #0 ;
 Value.G_N         := #0 ;
 Value.T_L         := #0 ;
 Value.T_N         := #0;
 Value.T_L         := #0 ;
 Value.G_V[1]      := #0 ;
 Value.G_V[2]      := #0 ;
 Value.G_V[3]      := #0 ;
 Value.G_V[4]      := #0 ;
 Value.G_V[5]      := #0 ;
 Value.CptCegid    := #0 ;
 Value.Dossier     := #0 ;
 Value.NatGene     := #0 ;
 Value.Etabliss    := #0 ;
 Value.Dev         := #0 ;
 Value.Exo         := #0 ;
end ;


end.
