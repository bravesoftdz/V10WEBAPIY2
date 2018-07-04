unit UtilSais;

interface

uses {$IFNDEF EAGLSERVER}
     Forms,
     Dialogs,
     Menus,
     Windows,
     Buttons,
     Grids,
     {$ENDIF}
     paramsoc,
     Classes,
     HCtrls,
     HQry,
     SysUtils,
     ExtCtrls,
     hent1,
    {$IFDEF EAGLCLIENT}
    {$ELSE}
    	DBGrids,DB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
      HDB,
    {$ENDIF}

    {$IFNDEF SANSCOMPTA}
			{$IFNDEF PGIIMMO}
  		  	HCompte,
			{$ENDIF}
		{$ENDIF}
     HMsgBox,
     Controls,
     ComCtrls,
     UTOB
     , ent1
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
  ,uEntCommun
     ;

Const SA_Exo      : Integer = 0 ;
      SA_NumP     : integer = 1 ;
      SA_DateC    : integer = 2 ;
      SA_NumL     : integer = 3 ;
      SA_Gen      : integer = 4 ;
      SA_Aux      : integer = 5 ;
      SA_RefI     : integer = 6 ;
      SA_Lib      : integer = 7 ;
      SA_Debit    : integer = 8 ;
      SA_Credit   : integer = 9 ;
      SA_ModeP    : Integer = 10 ;
      SA_DateEche : Integer = 11 ;
      SA_RIB      : Integer = 12 ;
      SA_RefE     : Integer = 13 ;
      SA_RefL     : Integer = 14 ;
      SA_NumTraChq: Integer = 15 ;
      SA_Jal      : integer = 0 ; // LG 24/02/2003
      SA_Etabl    : integer = 0 ; // LG 24/02/2003
      SA_Nat      : integer = 0 ; // LG 24/02/2003

      SA_PILESAI  : string  = '' ; // SBO empilement des saisies pour gestion des index de colonnes

Const AN_NumL     : Integer = 0 ;
      AN_Sect     : Integer = 1 ;
      AN_Lib      : Integer = 2 ;
      AN_Pourcent : Integer = 3 ;
      AN_Montant  : Integer = 4 ;
      AN_Date     : Integer = 5 ;

Const OA_Exo      : Integer = 0 ;
      OA_NumP     : integer = 1 ;
      OA_DateC    : integer = 2 ;
      OA_NumL     : integer = 3 ;
      {b FP FQ16213}
      OA_Axe      : Integer = 4 ;
      OA_Sect     : integer = 5 ;
      OA_RefI     : integer = 6 ;
      OA_Lib      : integer = 7 ;
      OA_Debit    : integer = 8 ;
      OA_Credit   : integer = 9 ;
      {e FP FQ16213}

Type TLettre = (NonEche,MultiEche,MonoEche) ;
Type tmSais = (tmNormal,tmPivot,tmDevise) ;

{$IFNDEF EAGLSERVER}
{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
PROCEDURE QUELZOOMTABLEG(Cpt : THCpteEdit ; Aux : TGTiers) ;
PROCEDURE QUELZOOMTABLET(Cpt : THCpteEdit ; Gen : TGGeneral) ;
function  Lettrable(CGen : TGGeneral ) : TLettre ;
Function  GetGGeneral ( GS : THGrid ; Lig : Longint ) : TGGeneral ;
Function  GetGTiers ( GS : THGrid ; Lig : Longint ) : TGTiers ;
Function  GetGSect ( GS : THGrid ; Lig : Longint ) : TGSection ;
Function  Ventilable ( CC : TObject ; i : integer) : Boolean ;
function  LigneEche ( GS : THGrid ; Lig : integer ) : boolean ;
procedure ChangeAffGrid (GS : THGrid ; Mode : tmSais ; Decim : byte) ;
Procedure InitDefautColSaisie ;
Function  FormatRIBEcr(Cpt : String ; AvecMess,VerifDansTable : Boolean ; Var St : String) : Integer ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
Procedure MajSoldesEcritureTOB ( TOBEcr : TOB ; Plus : boolean ; vInfo : TObject = nil ) ;
Function  QUELZOOMTABLETNAT ( NatureGene : String ) : TZoomTable ;
function  QUELDATATYPETNAT(vNatureGene: String) : String;
Function  CalculDateValeur(DateMvt : TDateTime ; t : TLettre ; Cpt : String) : TDateTime ;
FUNCTION  NATUREPIECECOMPTEOK (NatPiece,NatCpt : String3) : Boolean ;
FUNCTION  NATUREGENAUXOK (NatGen,NatAux : String3) : Boolean ;
FUNCTION  NATUREJALNATPIECEOK (const NatJal,NatPiece : String3) : Boolean ;
FUNCTION  CASENATP(Nat : String3) : Byte ;
FUNCTION  CASENATA(Nat : String3) : Byte ;
// LG*
procedure ExecMajSoldeTOB ( vTS : TList );
Procedure MajDesSoldesTOB ( TOBEcr : TOB ; Plus : boolean ) ;
procedure MajSoldeCompteTOB ( TOBEcr : TOB ; Plus : boolean ) ;
procedure MajSoldeSectionTOB ( TOBAna : TOB ; Plus : boolean; aDossier : string = '') ;
//procedure AjouteNew ( TS : TList; fb : TFichierBase ; Cpte : String17; D, C : double; NumP, NumL : Integer; DateP : tDateTime ; Axe : string3 = '' ; vInTypeExo : ShortInt = -1 ) ;
procedure AjouteTOB ( TS : TList;  vTOBligneEcr : TOB ; Plus : boolean ; vInTypeExo : ShortInt = -1 );
//SG6 16.02.05
procedure MajJournalAnaTob(TOBDetail : TOB ; Ano : boolean ;  Plus : boolean);

// MULTISOC
function  CUpdateCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vTotDeb, vTotCre, vTotDP, vTotCP, vTotDE, vTotCE, vTotDS, vTotCS : Double ; vBoAno : Boolean ; vDossier : String = '' ) : Boolean ;
procedure CChargeCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; var vTobCpt : TOB ; vDossier : String = '' ) ;
procedure CInsertCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vBoAll : Boolean = False) ;
procedure CDeleteCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vBoAll : Boolean ) ;
procedure CReinitCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vStCodeExo : String ; vBoODAnal : Boolean ; vBoAll : Boolean ) ;
procedure CMajTotCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vTotDeb, vTotCre : Double ; vTabDeb, vTabCre : Array Of Double ; vBoInit : Boolean ) ;
// SBO 19/04/2006 Ajout traitement sur totaux pointés pour les comptes généraux :
function  CUpdateCumulsPointeMS( vCompte : String ; vTotDebPTP, vTotCrePTP, vTotDebPTD, vTotCrePTD : Double ; vBoInit : Boolean = True ; vDossier : String = '' ) : Boolean ;
function  CReinitCumulsPointeMS( vCompte : String = '' ; vDossier : String = '' ) : Boolean ;
procedure CRAZCumulsMS( vFB : TFichierBase ; vAxe : string = '' ) ;
// SBO 20/04/2006 Ajout traitement pour insertion nouveau cumul :
function  CCreateCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vDossier : String = '' ) : Boolean ;

{$IFNDEF NOVH}
Procedure SauveSAEff ;
Procedure RestoreSAEff ;
Procedure initSA_SaisieNormale ;
{$ENDIF}
{$IFNDEF COMSX}
{$IFDEF COMPTA}
Function IsTiersSoumisTPF ( pGS : THGrid ) : Boolean ; // XVI 24/02/2005
{$ENDIF}
{$ENDIF}

procedure MajSoldeTOBAna ( TOBAna : TOB ; Plus : boolean; aDossier : string = '') ;

implementation

Uses SaisUtil,
     SaisComm,
{$IFNDEF PGIIMMO}
     ULibEcriture ,
{$ENDIF}
{$IFNDEF NOVH}
     uLibExercice, // CGetExerciceNMoins2
{$ENDIF NOVH}
{$IFDEF MODENT1}
     CPProcMetier,
{$ENDIF MODENT1}
     UtilPGI ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 22/11/2001
Modifié le ... :   /  /
Description .. : Enregistrement en base des soldes des comptes,des tiers...
Mots clefs ... : MAJ COMPTE
*****************************************************************}
procedure ExecMajSoldeTOB ( vTS : TList );
var
 pT : ^TFRM;


 procedure ExecMAJ( fb : TFichierBase ; vTS : TList);
 var
  i  : integer;
 begin
  for i := 0 to vTS.Count - 1 do
  begin
   pT := vTS.Items[i];
   if ( pT^.fb = fb ) then
    begin
     if ExecReqMAJ ( fb, false , false, pT^ ) <> 1 then
      begin
       V_PGI.IoError := oeSaisie;
       exit;
      end; // if
   end; // if
  end; // for
 end;


begin

  ExecMAJ (fbGene,vTS);  // mises à jour des generaux
  if V_PGI.IoError=oeOk then ExecMAJ ( fbAux , vTS );  // maj des aux
  if V_PGI.IoError=oeOk then ExecMAJ ( fbSect , vTS ); // maj des sections
  if V_PGI.IoError=oeOk then ExecMAJ ( fbJal , vTS );  // maj des journaux

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 13/11/2001
Modifié le ... :   /  /
Description .. : Ajoute ou cree les elements a la TList TS
Mots clefs ... : MAJ COMPTE
*****************************************************************}
procedure AjouteNew ( TS : TList;fb : TFichierBase ; Cpte : String17; D, C : double; NumP, NumL : Integer; DateP : TDateTime ; Axe : string3 = '' ; vInTypeExo : ShortInt = -1 ) ;
var
 i        : integer ;
 lpT      : ^TFRM ;
 lTypeExo : TTypeExo ;
begin
  if Trim ( Cpte ) = '' then Exit;

  if vInTypeExo <> -1
    then lTypeExo := TTypeExo( vInTypeExo )
    else lTypeExo := CGetTypeExo(DateP) ;

  for i := 0 to TS.Count - 1 do
   begin
   lpT := TS.Items[i];
   if (lpT^.fb=fb) and (lpT^.Cpt=Cpte) and (lpT^.Axe=Axe) then
     begin
     Case lTypeExo of
       teEncours : BEGIN
                   lpT^.DE := lpT^.DE+D ;
                   lpT^.CE := lpT^.CE+C ;
                   END ;
       teSuivant : BEGIN
                   lpT^.DS := lpT^.DS+D ;
                   lpT^.CS := lpT^.CS+C ;
                   END ;
       tePrecedent : BEGIN
                     lpT^.DP := lpT^.DP+D ;
                     lpT^.CP := lpT^.CP+C ;
                     END ;
       END ;
     if DateP >= lpT^.DateD then
       begin
       lpT^.Deb   := D;
       lpT^.Cre   := C;
       lpT^.DateD := DateP;
       lpT^.NumD  := NumP ;
       lpT^.LigD  := NumL;
       end; // if
     Exit;
     end;
  end;
 New(lpT);
 FillChar(lpT^,sizeOf(lpT^),#0);
 AttribParamsNew( lpT^, D, C, lTypeExo );
 lpT^.fb    := fb ;
 lpT^.Cpt   := Cpte ;
 lpT^.Axe   := Axe ;
 lpT^.Deb   := D ;
 lpT^.Cre   := C ;
 lpT^.DateD := DateP ;
 lpT^.NumD  := NumP ;
 lpT^.LigD  := NumL ;
 TS.Add(lpT);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 13/11/2001
Modifié le ... : 16/10/2003
Description .. : Cumul les debit et des credit du compte, du tiers et des
Suite ........ : sections de la ligne d'ecriture vTOBLigneEcr et les ajoute à
Suite ........ : la TList TS.
Suite ........ : - 19/03/2002 : correction : if plus then remplacer par if not
Suite ........ : plus
Suite ........ : - 16/10/2003 - meme correction pour les sections
Mots clefs ... : MAJ COMPTE
*****************************************************************}
procedure AjouteTOB ( TS : TList;  vTOBLigneEcr : TOB ; Plus : boolean ; vInTypeExo : ShortInt = -1 );
var
 j,l           : integer;
 lDateC        : TDateTime; // Date comptable
 lLig          : integer; // numero de ligne
 lNumP         : Integer; // Numero de piece
 lInCpt        : Integer ; //Compteur
 lXD           : double; // Debit
 lXC           : double; // Credit
 lAx           : string; // Axe analytique
 lSect         : string; // Section analytique
 lGen          : string; // compte general
 lAux          : string; // compte auxiliaire
 lJrn          : string; // code du journal
begin
lGen:=vTOBLigneEcr.GetValue('E_GENERAL'); lAux:=vTOBLigneEcr.GetValue('E_AUXILIAIRE');
lXD:=vTOBLigneEcr.GetValue('E_DEBIT'); lXC:=vTOBLigneEcr.GetValue('E_CREDIT');
lNumP:=vTOBLigneEcr.GetValue('E_NUMEROPIECE'); lLig:=vTOBLigneEcr.GetValue('E_NUMLIGNE') ;
lDateC:=vTOBLigneEcr.GetValue('E_DATECOMPTABLE') ; lJrn:=vTOBLigneEcr.GetValue('E_JOURNAL') ;
if not Plus then begin lXD := lXD * (-1); lXC := lXC * (-1); end; // if
// ajoute ou creer les valeurs à la TList lTS
AjouteNew( TS, fbGene , lGen , lXD, lXC , lNumP, lLig , lDateC, '', vInTypeExo ); // ajout au generaux
AjouteNew( TS, fbAux  , lAux , lXD, lXC , lNumP, lLig , lDateC, '', vInTypeExo ); // ajout des tiers
AjouteNew( TS, fbJal  , lJrn , lXD, lXC , lNumP, lLig,  lDateC, '', vInTypeExo ); // ajout du journal
// parcourd des lignes d'analytiques
if vTOBLigneEcr.GetValue('E_ANA') = 'X' then
 begin
   for l := 0 to vTOBLigneEcr.Detail.Count - 1 do
    for j := 0 to vTOBLigneEcr.Detail[l].Detail.Count - 1 do
     begin
      lXD     := vTOBLigneEcr.Detail[l].Detail[j].GetValue('Y_DEBIT') ;
      lXC     := vTOBLigneEcr.Detail[l].Detail[j].GetValue('Y_CREDIT') ;
      if not Plus then begin lXD:=lXD*(-1); lXC:=lXC*(-1); end; // if

      //SG6 27.01.05 Gestion mode croisaxe
      if not (GetAnaCroisaxe) then
      begin
        lSect   := vTOBLigneEcr.Detail[l].Detail[j].GetValue('Y_SECTION');
        lAx     := vTOBLigneEcr.Detail[l].Detail[j].GetValue('Y_AXE') ;
        AjouteNew( TS, fbSect , lSect , lXD , lXC, lNumP, lLig, lDateC, lAx, vInTypeExo ) ;
      end
      else
      begin
        for lInCpt := 1 to MaxAxe do
        begin
          lSect   := vTOBLigneEcr.Detail[l].Detail[j].GetValue('Y_SOUSPLAN' + IntToStr( lInCpt ));
          lAx     := 'A' + IntToStr(lInCpt);
          if lSect <> '' then
          begin
            AjouteNew( TS, fbSect , lSect , lXD , lXC, lNumP, lLig, lDateC, lAx, vInTypeExo ) ;
          end;
        end;
      end;


     end; // for
 end; // if Analytique
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 22/11/2001
Modifié le ... : 02/06/2006
Description .. : Fonction de mises à jour des soldes des comptes, des tiers,
Suite ........ : des section et de l'analytiques.
Suite ........ : Param :
Suite ........ :  - TOBEcr : TOB ecriture
Suite ........ :  - Plus : si true les montants sont additionné
Suite ........ : - LG - 10/05/2004 - ajout de la devalidation du journal
Suite ........ : - FB 10678 - LG - 02/06/2006 - gestion des devises ds les 
Suite ........ : ano dynamiques
Mots clefs ... : MAJ COMPTE
*****************************************************************}
Procedure MajSoldesEcritureTOB ( TOBEcr : TOB ; Plus : boolean ; vInfo : TObject = nil ) ;
Var
 i               : integer;
 lTS             : TList;
 lTSA            : TList ;
 lTOBLigneEcr    : TOB;
 lStJal          : string ;
 lDtDateC        : TDateTime ;
{$IFNDEF PGIIMMO}
 lInfo           : TInfoEcriture ;
{$ENDIF}
 lBoDetruireInfo : boolean ;
 pT              : ^TFRM ;
BEGIN
 // création de la TList contenant les TFRM
 lTS             := TList.Create ;
 lTSA            := TList.Create ;
 lDtDateC        := 0 ;
 lStJal          := '' ;
 lBoDetruireInfo := false ;

{$IFNDEF PGIIMMO}
 if ( vInfo <> nil ) and ( vInfo is TInfoEcriture ) then
  lInfo := TInfoEcriture(vInfo)
   else
    begin
     lInfo := TInfoEcriture.Create( V_PGI.SchemaName ) ;
     lBoDetruireInfo := true ;
    end ;// if
{$ENDIF}

 try

 for i := 0 to TOBEcr.Detail.Count - 1 do
  begin
   lTOBLigneEcr := TOBEcr.Detail[i];
   lStJal       := lTOBLigneEcr.GetValue('E_JOURNAL') ;
   lDtDateC     := lTOBLigneEcr.GetValue('E_DATECOMPTABLE') ;
   AjouteTOB(lTS,  lTOBLigneEcr , Plus ); // ajoute les elements à la TList
{$IFNDEF PGIIMMO}
   if _GetParamSocSecur('SO_CPANODYNA',false) then
    begin
     lInfo.LoadCompte(lTOBLigneEcr.GetValue('E_GENERAL')) ;
     AjouteAno(lTSA,lTOBLigneEcr,lInfo.GetString('G_NATUREGENE'), not Plus) ;
    end ;
{$ENDIF}
  end; // if

 // Gestion MultiSoc
 for i := 0 to lTS.Count-1 do
  begin
   pT := lTS.Items[ i ] ;
{$IFNDEF PGIIMMO}
   pT.Dossier := lInfo.Dossier ;
{$ENDIF}
  end ;

 // enregistre en base
 ExecMajSoldeTOB (lTS) ;
 ADevalider(lStJal,lDtDateC) ;

 if _GetParamSocSecur('SO_CPANODYNA',false) then
  begin
   if not ExecReqMAJAno(lTSA) then
     begin
      V_PGI.IoError := oeSaisie ;
      exit ;
     end ;
  end ;

 finally
  if lBoDetruireInfo then
{$IFNDEF PGIIMMO}
   lInfo.Free ;
{$ENDIF}

  for i:=0 to lTS.Count-1 do
   begin
    pT := lTS[i] ;
    Dispose(pT) ;
    lTS[i] := nil ;
   end ;

  for i:=0 to lTSA.Count-1 do
   begin
    pT := lTSA[i] ;
    Dispose(pT) ;
    lTSA[i] := nil ;
   end ;

  lTS.Free ;
  lTSA.Free ;
  
 end;

END ;

Procedure MajDesSoldesTOB ( TOBEcr : TOB ; Plus : boolean ) ;
Var i,j : integer ;
BEGIN
MajSoldeCompteTOB(TOBEcr,Plus) ;
if TOBEcr.GetValue('E_ANA')='X' then
   BEGIN
   for i:=0 to TOBEcr.Detail.Count-1 do
       for j:=0 to TOBEcr.Detail[i].Detail.Count-1 do MajSoldeSectionTOB(TOBEcr.Detail[i].Detail[j],Plus) ;
   END ;
END ;


procedure MajSoldeCompteTOB ( TOBEcr : TOB ; Plus : boolean ) ;
Var StE   : String3 ;
    DateC : TDateTime ;
    Num   : Longint ;
    Lig   : integer ;
    CG,CA,Jal,SQL : String ;
    XD,XC : Double ;
    DE,CE,DP,CP,DS,CS : Double ;
BEGIN
{capture des infos en vue des MAJ}
DateC:=TOBEcr.GetValue('E_DATECOMPTABLE') ; StE:=QuelExo(DateToStr(DATEC)) ;
Num:=TOBEcr.GetValue('E_NUMEROPIECE') ; Lig:=TOBEcr.GetValue('E_NUMLIGNE') ;
CG:=TOBEcr.GetValue('E_GENERAL') ; CA:=TOBEcr.GetValue('E_AUXILIAIRE') ;
XD:=TOBEcr.GetValue('E_DEBIT') ; XC:=TOBEcr.GetValue('E_CREDIT') ;
if Not Plus then BEGIN XD:=-XD ; XC:=-XC ; END ;
Jal:=TOBEcr.GetValue('E_JOURNAL') ;
DE:=0 ; CE:=0 ; DP:=0 ; CP:=0 ; DS:=0 ; CS:=0 ;
if StE=GetEncours.Code then BEGIN DE:=XD ; CE:=XC; END else
 if StE=GetSuivant.Code then BEGIN DS:=XD ; CS:=XC ; END else BEGIN DP:=XC ; CP:=XC ; END ;
{Général MAJ}
if not EstTablePartagee( 'GENERAUX' ) then
  begin
  SQL:='UPDATE GENERAUX SET G_DEBITDERNMVT='+StrfPoint(XD)+',  G_CREDITDERNMVT='+StrfPoint(XC)+', '
      +'G_DATEDERNMVT="'+USDateTime(DateC)+'" , G_NUMDERNMVT='+IntToStr(Num)+', G_LIGNEDERNMVT='+IntToStr(Lig)+', '
      +'G_TOTALDEBIT=G_TOTALDEBIT+'+StrfPoint(XD)+', G_TOTALCREDIT=G_TOTALCREDIT+'+StrfPoint(XC)+', '
      +'G_TOTDEBE=G_TOTDEBE+'+StrfPoint(DE)+', G_TOTCREE=G_TOTCREE+'+StrfPoint(CE)+', '
      +'G_TOTDEBS=G_TOTDEBS+'+StrfPoint(DS)+', G_TOTCRES=G_TOTCRES+'+StrfPoint(CS)+', '
      +'G_TOTDEBP=G_TOTDEBP+'+StrfPoint(DP)+', G_TOTCREP=G_TOTCREP+'+StrfPoint(CP)+' WHERE G_GENERAL="'+CG+'"' ;
    if ExecuteSQL(SQL)<>1
      then V_PGI.IoError:= oeLettrage ; // FQ 20001 BVE 25.04.07 On utilise oeLettrage car pas de oeGeneral
  end
else CUpdateCumulsMS( fbGene, CG, '', XD, XC, DP, CP, DE, CE, DS, CS, False ) ;
{Auxi MAJ}
if ((CA<>'') and (V_PGI.IoError=oeOk)) then
   BEGIN
   if not EstTablePartagee( 'TIERS' ) then
     begin
     SQL:='UPDATE TIERS SET T_DEBITDERNMVT='+StrfPoint(XD)+',  T_CREDITDERNMVT='+StrfPoint(XC)+', '
         +'T_DATEDERNMVT="'+USDateTime(DateC)+'" , T_NUMDERNMVT='+IntToStr(Num)+', T_LIGNEDERNMVT='+IntToStr(Lig)+', '
         +'T_TOTALDEBIT=T_TOTALDEBIT+'+StrfPoint(XD)+', T_TOTALCREDIT=T_TOTALCREDIT+'+StrfPoint(XC)+', '
         +'T_TOTDEBE=T_TOTDEBE+'+StrfPoint(DE)+', T_TOTCREE=T_TOTCREE+'+StrfPoint(CE)+', '
         +'T_TOTDEBS=T_TOTDEBS+'+StrfPoint(DS)+', T_TOTCRES=T_TOTCRES+'+StrfPoint(CS)+', '
         +'T_TOTDEBP=T_TOTDEBP+'+StrfPoint(DP)+', T_TOTCREP=T_TOTCREP+'+StrfPoint(CP)+' WHERE T_AUXILIAIRE="'+CA+'"' ;
     if ExecuteSQL(SQL)<>1
       then V_PGI.IoError:=oeTiers ; // FQ 20001 BVE 25.04.07
     end
   else CUpdateCumulsMS( fbAux, CA, '', XD, XC, DP, CP, DE, CE, DS, CS, False ) ;
   END ;
{Journal MAJ}
if V_PGI.IoError=oeOk then
   BEGIN
   if not EstTablePartagee( 'JOURNAL' ) then
     begin
     SQL:='UPDATE JOURNAL SET J_DEBITDERNMVT='+StrfPoint(XD)+', J_CREDITDERNMVT='+StrfPoint(XC)+', '
         +'J_DATEDERNMVT="'+USDateTime(DateC)+'", J_NUMDERNMVT='+IntToStr(Num)+', J_TOTALDEBIT=J_TOTALDEBIT+'+Strfpoint(XD)+', '
         +'J_TOTALCREDIT=J_TOTALCREDIT+'+StrfPoint(XC)+', J_TOTDEBE=J_TOTDEBE+'+StrfPoint(DE)+', J_TOTCREE=J_TOTCREE+'+StrfPoint(CE)+','
         +'J_TOTDEBS=J_TOTDEBS+'+StrfPoint(DS)+', J_TOTCRES=J_TOTCRES+'+StrfPoint(CS)+', '
         +'J_TOTDEBP=J_TOTDEBP+'+StrfPoint(DP)+', J_TOTCREP=J_TOTCREP+'+StrfPoint(CP)+' WHERE J_JOURNAL="'+Jal+'"' ;
     if ExecuteSQL(SQL)<>1
       then V_PGI.IoError:=oeStock;  // FQ 20001 BVE 25.04.07 On utilise oeStock car pas de oeJournal
     end
   else CUpdateCumulsMS( fbJal, Jal, '', XD, XC, DP, CP, DE, CE, DS, CS, False ) ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 16/02/2005
Modifié le ... :   /  /
Description .. : Procedure qui permet de mettre a jour les soldes
Suite ........ : des journaux pour les ecriture d'ODA.
Mots clefs ... : MAJ SOLDE JAL ODA
*****************************************************************}
procedure MajJournalAnaTob(TOBDetail : TOB ; Ano : boolean ;  Plus : boolean);
var lFRM        : TFRM;
    GeneTypeExo : TTypeExo;
    XD, XC      : Double;
begin

  GeneTypeExo := CGetTypeExo( TOBDetail.GetDateTime('Y_DATECOMPTABLE') ) ;

  XD := TOBDetail.GetValue('Y_DEBIT');
  XC := TOBDetail.GetValue('Y_CREDIT');

  lFRM.Cpt := TOBDetail.GetValue('Y_JOURNAL');

  if not Plus then
  begin
    XD := -XD;
    XC := -XC;
  end;

  lFRM.Axe := TOBDetail.GetValue('Y_AXE');
  if not Ano then
  begin
    lFRM.NumD := TOBDetail.GetValue('Y_NUMEROPIECE');
    lFRM.DateD := TOBDetail.GetValue('Y_DATECOMPTABLE');
    AttribParamsNew(lFRM, XD, XC, GeneTypeExo);
  end
  else
  begin
    lFRM.Deb := XD;
    lFRM.Cre := XC;
  end;

  if ExecReqMAJ(fbJal, Ano, False, lFRM) <> 1 then
        V_PGI.IoError := oeSaisie;
end;


procedure MajSoldeJournalTOB ( TOBAna : TOB ; Plus : boolean; aDossier : string = '') ;
Var XD,XC : double ;
    StE,SQL,Ax,Section,Jal : String ;
    Num,Lig, lInCpt : integer ;
    DateC   : TDateTime ;
    DE,CE,DP,CP,DS,CS : Double ;
BEGIN

  XD := TOBAna.GetValue('Y_DEBIT') ;
  XC := TOBAna.GetValue('Y_CREDIT') ;
  if Not Plus then begin XD:=-XD ; XC:=-XC ; end ;

  Jal := TOBAna.GetValue('Y_JOURNAL') ;
  Num := TOBAna.GetValue('Y_NUMEROPIECE') ;
  Lig := TOBAna.GetValue('Y_NUMLIGNE') ;
  Ax := TOBAna.GetValue('Y_AXE') ;
  Section := TOBAna.GetValue('Y_SECTION') ;
  DateC:=TOBAna.GetValue('Y_DATECOMPTABLE') ;
  StE:=TOBAna.GetValue('Y_EXERCICE') ;
  DE := 0 ; CE := 0 ; DP := 0 ; CP := 0 ; DS := 0 ; CS := 0 ;
  if StE=GetEncours.Code then begin DE := XD ; CE := XC; end else
  if StE=GetSuivant.Code then begin DS:=XD ; CS:=XC ; end else BEGIN DP := XC ; CP := XC ; end ;


  if not EstTablePartagee( 'JOURNAL' ) then
  begin
    SQL:='UPDATE JOURNAL SET J_DEBITDERNMVT='+StrfPoint(XD)+', J_CREDITDERNMVT='+StrfPoint(XC)+', '
          +'J_DATEDERNMVT="'+USDateTime(DateC)+'", J_NUMDERNMVT='+IntToStr(Num)+', J_TOTALDEBIT=J_TOTALDEBIT+'+Strfpoint(XD)+', '
          +'J_TOTALCREDIT=J_TOTALCREDIT+'+StrfPoint(XC)+', J_TOTDEBE=J_TOTDEBE+'+StrfPoint(DE)+', J_TOTCREE=J_TOTCREE+'+StrfPoint(CE)+','
          +'J_TOTDEBS=J_TOTDEBS+'+StrfPoint(DS)+', J_TOTCRES=J_TOTCRES+'+StrfPoint(CS)+', '
          +'J_TOTDEBP=J_TOTDEBP+'+StrfPoint(DP)+', J_TOTCREP=J_TOTCREP+'+StrfPoint(CP)+' WHERE J_JOURNAL="'+Jal+'"' ;
    if ExecuteSQL(SQL)<>1
    then V_PGI.IoError:=oeStock;  // FQ 20001 BVE 25.04.07 On utilise oeStock car pas de oeJournal
  end
  else CUpdateCumulsMS( fbJal, Jal, '', XD, XC, DP, CP, DE, CE, DS, CS, False ) ;
end;


procedure MajSoldeTOBAna ( TOBAna : TOB ; Plus : boolean; aDossier : string = '') ;
begin
  MajSoldeSectionTOB (TOBAna,Plus,aDossier);
  if V_PGI.IoError <> oeOk then Exit;
  MajSoldeJournalTOB ( TOBAna,Plus,aDossier);
end;


procedure MajSoldeSectionTOB ( TOBAna : TOB ; Plus : boolean; aDossier : string = '') ;
Var XD,XC : double ;
    StE,SQL,Ax,Section : String ;
    Num,Lig, lInCpt : integer ;
    DateC   : TDateTime ;
    DE,CE,DP,CP,DS,CS : Double ;
BEGIN
XD := TOBAna.GetValue('Y_DEBIT') ;
XC := TOBAna.GetValue('Y_CREDIT') ;
if Not Plus then begin XD:=-XD ; XC:=-XC ; end ;

Num := TOBAna.GetValue('Y_NUMEROPIECE') ;
Lig := TOBAna.GetValue('Y_NUMLIGNE') ;
Ax := TOBAna.GetValue('Y_AXE') ;
Section := TOBAna.GetValue('Y_SECTION') ;
DateC:=TOBAna.GetValue('Y_DATECOMPTABLE') ;
StE:=TOBAna.GetValue('Y_EXERCICE') ;
DE := 0 ; CE := 0 ; DP := 0 ; CP := 0 ; DS := 0 ; CS := 0 ;
if StE=GetEncours.Code then begin DE := XD ; CE := XC; end else
 if StE=GetSuivant.Code then begin DS:=XD ; CS:=XC ; end else BEGIN DP := XC ; CP := XC ; end ;

//SG6 27.01.05 Gestion mode croisaxe

if not(GetAnaCroisaxe) then
begin
  if VH^.LiaisonY2ViaShare then
  begin
    SQL:='UPDATE CSECTION SET CSP_DEBITDERNMVT='+StrfPoint(XD)+',  CSP_CREDITDERNMVT='+StrfPoint(XC)+', '
      +'CSP_DATEDERNMVT="'+UsDateTime(DateC)+'", CSP_NUMDERNMVT='+IntToStr(Num)+', CSP_LIGNEDERNMVT='+IntToStr(Lig)+', '
      +'CSP_TOTALDEBIT=CSP_TOTALDEBIT+'+StrfPoint(XD)+', CSP_TOTALCREDIT=CSP_TOTALCREDIT+'+StrfPoint(XC)+', '
      +'CSP_TOTDEBE=S_TOTDEBE+'+StrfPoint(DE)+', CSP_TOTCREE=CSP_TOTCREE+'+StrfPoint(CE)+', '
      +'CSP_TOTDEBS=S_TOTDEBS+'+Strfpoint(DS)+', CSP_TOTCRES=CSP_TOTCRES+'+StrfPoint(CS)+', '
      +'CSP_TOTDEBP=CSP_TOTDEBP+'+StrfPoint(DP)+', CSP_TOTCREP=CSP_TOTCREP+'+StrfPoint(CP)+' WHERE CSP_AXE="'+AX+'" AND CSP_SECTION="'+Section+'"' ;
    if ExecuteSQL(SQL)<>1 then V_PGI.IoError:=oeSaisie ;
  end else
  begin
    if not EstTablePartagee( 'SECTION' ) then
      begin
      SQL:='UPDATE SECTION SET S_DEBITDERNMVT='+StrfPoint(XD)+',  S_CREDITDERNMVT='+StrfPoint(XC)+', '
        +'S_DATEDERNMVT="'+UsDateTime(DateC)+'", S_NUMDERNMVT='+IntToStr(Num)+', S_LIGNEDERNMVT='+IntToStr(Lig)+', '
        +'S_TOTALDEBIT=S_TOTALDEBIT+'+StrfPoint(XD)+', S_TOTALCREDIT=S_TOTALCREDIT+'+StrfPoint(XC)+', '
        +'S_TOTDEBE=S_TOTDEBE+'+StrfPoint(DE)+', S_TOTCREE=S_TOTCREE+'+StrfPoint(CE)+', '
        +'S_TOTDEBS=S_TOTDEBS+'+Strfpoint(DS)+', S_TOTCRES=S_TOTCRES+'+StrfPoint(CS)+', '
        +'S_TOTDEBP=S_TOTDEBP+'+StrfPoint(DP)+', S_TOTCREP=S_TOTCREP+'+StrfPoint(CP)+' WHERE S_AXE="'+AX+'" AND S_SECTION="'+Section+'"' ;
      if ExecuteSQL(SQL)<>1
        then V_PGI.IoError:=oeSaisie ;
      end
    else CUpdateCumulsMS( fbSect, Section, Ax, XD, XC, DP, CP, DE, CE, DS, CS, False, aDossier ) ;
  end;
end
else
begin
  for lInCpt := 1 to MaxAxe do
  begin
    Section := TOBAna.GetString( 'Y_SOUSPLAN' + IntToStr(lInCpt)) ;
    AX := 'A' + IntToStr( lInCpt ) ;
    if Section <> '' then
    begin
      if not EstTablePartagee( 'SECTION' ) then
        begin
        SQL:='UPDATE SECTION SET S_DEBITDERNMVT='+StrfPoint(XD)+',  S_CREDITDERNMVT='+StrfPoint(XC)+', '
          +'S_DATEDERNMVT="'+UsDateTime(DateC)+'", S_NUMDERNMVT='+IntToStr(Num)+', S_LIGNEDERNMVT='+IntToStr(Lig)+', '
          +'S_TOTALDEBIT=S_TOTALDEBIT+'+StrfPoint(XD)+', S_TOTALCREDIT=S_TOTALCREDIT+'+StrfPoint(XC)+', '
          +'S_TOTDEBE=S_TOTDEBE+'+StrfPoint(DE)+', S_TOTCREE=S_TOTCREE+'+StrfPoint(CE)+', '
          +'S_TOTDEBS=S_TOTDEBS+'+Strfpoint(DS)+', S_TOTCRES=S_TOTCRES+'+StrfPoint(CS)+', '
          +'S_TOTDEBP=S_TOTDEBP+'+StrfPoint(DP)+', S_TOTCREP=S_TOTCREP+'+StrfPoint(CP)+' WHERE S_AXE="'+ AX +'" AND S_SECTION="'+ Section +'"' ;
        if ExecuteSQL(SQL)<>1
          then V_PGI.IoError:=oeSaisie ;
        end
      else CUpdateCumulsMS( fbSect, Section, Ax, XD, XC, DP, CP, DE, CE, DS, CS, False, aDossier ) ;
    end;
  end;
end;
END ;

//LG*

FUNCTION CASENATC(Nat : String3) : Byte ;
begin { Pour les généraux }
Result:=9 ; { = 'DIV' }
If Nat='COC' then Result:=1  else if Nat='COF' then Result:=2 else if Nat='BQE' then Result:=3  else
If Nat='CAI' then Result:=4  else if Nat='IMO' then Result:=5 else if Nat='COS' then Result:=6  else
If Nat='CHA' then Result:=7  else if Nat='PRO' then Result:=8 else if Nat='EXT' then Result:=10 else
If Nat='TID' then Result:=11 else if Nat='TIC' then Result:=12 ;
end ;

FUNCTION CASENATA(Nat : String3) : Byte ;
begin { Pour les tiers }
Result:=9 ; { = 'DIV' }
If (Nat='AUD') Or (Nat='CLI') then Result:=1  else
If (Nat='AUC') Or (Nat='FOU') then Result:=2  else
if Nat='SAL' then Result:=6 ;
end ;


Function QUELZOOMTABLETNAT ( NatureGene : String ) : TZoomTable ;
Var INatGen : Byte ;
BEGIN
INatGen:=CaseNatC(NatureGene) ;
Case INatGen Of
  1 : Result:=tzTToutDebit ;
  2 : Result:=tzTToutCredit ;
  6 : result:=tzTSalarie ;
  else result:=tzTiers ;
  END ;
END ;

function QUELDATATYPETNAT(vNatureGene : String) : String;
var lNatGen : Byte ;
begin
  Result := 'TZTTous' ;
  lNatGen := CASENATC( vNatureGene );
  Case lNatGen Of
    1 : Result  := 'TZTToutDebit' ;
    2 : Result  := 'TZTToutCredit' ;
    6 : result  := 'TZTSalarie' ;
    else result := 'TZTTous' ;
  end;
end;

{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
{$IFNDEF EAGLSERVER}
PROCEDURE QUELZOOMTABLEG(Cpt : THCpteEdit ; Aux : TGTiers) ;
Var INatAux : Byte ;
begin
If Aux=Nil Then Exit ; If Cpt=Nil Then Exit ;
INatAux:=CaseNatA(Aux.NatureAux) ;
Case INatAux Of
  1 : Cpt.ZoomTable:=tzGCollToutDebit ;
  2 : Cpt.ZoomTable:=tzGCollToutCredit ;
  6 : Cpt.ZoomTable:=tzGCollSalarie ;
  Else Cpt.ZoomTable:=tzGCollectif ;
  END ;
end ;

PROCEDURE QUELZOOMTABLET(Cpt : THCpteEdit ; Gen : TGGeneral) ;
begin
If Gen=Nil Then Exit ; If Cpt=Nil Then Exit ;
Cpt.ZoomTable:=QuelZoomTableTNat(Gen.NatureGene) ;
end ;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : ???
Créé le ...... : 01/01/1900
Modifié le ... : 08/08/2005
Description .. : 
Suite ........ : SBO 08/08/2005 : Adaptation aux comptes "divers" 
Suite ........ : "lettrable"
Mots clefs ... : 
*****************************************************************}
function Lettrable(CGen : TGGeneral ) : TLettre ;
begin
  Result:=NonEche ;
  if CGen=Nil then exit ;
  if (CGen.Lettrable) then
    if CGen.NatureGene = 'DIV'
      then Result := MonoEche
      else Result := MultiEche
  else if CGen.Pointable and ((CGen.NatureGene='BQE') or (CGen.NatureGene='CAI')) then
    Result := MonoEche ;
end ;

Function Ventilable ( CC : TObject ; i : integer) : Boolean ;
Var j : integer ;
    VV : T5B ;
BEGIN
Result:=False ; if CC=Nil then Exit ;
if CC is TGGeneral then VV:=TGGeneral(CC).Ventilable else Exit ;
if i>0 then Result:=VV[i] else
   BEGIN
   Result:=False ;
   for j:=1 to MaxAxe do Result:=((Result) or (VV[j])) ;
   END ;
END ;

{$IFNDEF EAGLSERVER}
function LigneEche ( GS : THGrid ; Lig : integer ) : boolean ;
Var CGen : TGGeneral ;
    CAux : TGTiers ;
    t    : TLettre ;
BEGIN
Result:=False ;
CGen:=GetGGeneral(GS,Lig) ; if CGen=Nil then exit ;
CAux:=GetGTiers(GS,Lig) ;
if CGen.Collectif then
   BEGIN
   if CAux=Nil then Exit ; if Not CAux.Lettrable then Exit ;
   Result:=True ;
   END else
   BEGIN
   t:=Lettrable(CGen) ; if t<>NonEche then Result:=True ;
   END ;
END ;

Function GetGGeneral ( GS : THGrid ; Lig : Longint ) : TGGeneral ;
BEGIN Result:=TGGeneral(GS.Objects[SA_Gen,Lig]) ; END ;

Function GetGTiers ( GS : THGrid ; Lig : Longint ) : TGTiers ;
BEGIN Result:=TGTiers(GS.Objects[SA_Aux,Lig]) ; END ;

Function GetGSect ( GS : THGrid ; Lig : Longint ) : TGSection ;
BEGIN
if GS.TypeSais=tsODA then Result:=TGSection(GS.Objects[OA_Sect,Lig]) else Result:=TGSection(GS.Objects[AN_Sect,Lig]) ;
END ;
{$ENDIF}
{$ENDIF}
{$ENDIF}

Function CalculDateValeur(DateMvt : TDateTime ; t : TLettre ; Cpt : String) : TDateTime ;
begin
Result:=DateMvt ;
end ;

FUNCTION NATUREPIECECOMPTEOK (NatPiece,NatCpt : String3) : Boolean ;
Var INatPiece,INatCpt : Byte ;
begin
Result:=FALSE  ;
INatPiece:=CaseNatP(NatPiece) ; INatCpt:=CaseNatC(NatCpt) ;
(*
Case INatPiece of
  1,2,3 : if INatCpt in [2,12] then Exit ;
  4,5,6 : if INatCpt in [1,11] then Exit ;
  end ;
*)
if ctxPCL in V_PGI.PGIContexte then
  BEGIN
  Case INatPiece of
    1,2,3 : if INatCpt in [2] then Exit ;
    4,5,6 : if INatCpt in [1] then Exit ;
    end ;
  END Else
  BEGIN
(* Modif GP le 03/06/2002
{$IFDEF CEGID}
  Case INatPiece of
    1,2,3 : if INatCpt in [2] then Exit ;
    4,5,6 : if INatCpt in [1] then Exit ;
    end ;
{$ELSE}
  Case INatPiece of
    1,2,3 : if INatCpt in [2,12] then Exit ;
    4,5,6 : if INatCpt in [1,11] then Exit ;
    end ;
{$ENDIF}
*)
  Case INatPiece of
    1,2,3 : if INatCpt in [2] then Exit ;
    4,5,6 : if INatCpt in [1] then Exit ;
    end ;
  END ;
Result:=TRUE  ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 18/03/2003
Modifié le ... :   /  /    
Description .. : - 18/03/2003 - correction pour la version sic. on peut saisir 
Suite ........ : des collectifs sur des comptes de charges
Mots clefs ... : 
*****************************************************************}
FUNCTION NATUREGENAUXOK (NatGen,NatAux : String3) : Boolean ;
Var INatGen,INatAux : Byte ;
begin
if GetCPIFDEFCEGID then
Begin
Result:=FALSE  ;
INatAux:=CaseNatA(NatAux) ; INatGen:=CaseNatC(NatGen) ;
Case INatGen of
  1 :  if (INatAux in [1,9])=FALSE then Exit ;
  2  : if (INatAux in [2,9])=FALSE then Exit ;
  end ;
Result:=TRUE  ;
end else begin
result:=    ((NatGen='COF') and (NatAux='FOU'))
         or ((NatGen='COF') and (NatAux='AUC')) // FQ 12769
         or ((NatGen='COF') and (NatAux='DIV')) // FQ 12769
         or (NatGen='COD') // FQ 12769
         or ((NatGen='COC') and (NatAux='CLI'))
         or ((NatGen='COC') and (NatAux='AUD')) // FQ 12769
         or ((NatGen='COC') and (NatAux='DIV')) // FQ 12769
         or ((NatGen='COS') and (NatAux='SAL'))
         or ((NatGen='DIV') and (NatAux='DIV'));
end ;
end ;

FUNCTION NATUREJALNATPIECEOK (const NatJal,NatPiece : String3) : Boolean ;
var INatPiece : Byte ;
begin
 INatPiece:=CASENATP(NatPiece) ;
 case CaseNatJal(NatJal) of
  tzJVente       : result:=(INatPiece in [1,2,7]) ;
  tzJAchat       : result:=(INatPiece in [4,5,7]) ;
  tzJBanque      : result:=(INatPiece in [3,6,7]) ;
  tzJEcartChange : result:= (INatPiece in [7,8]) ; 
  tzJOD          : result:=true;
  else
   result:=false;
 end; // case
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 09/10/2002
Modifié le ... :   /  /    
Description .. : - 09/10/2002 - Ajout des ecritures d'ecart de conversion 
Suite ........ : euro
Mots clefs ... : 
*****************************************************************}
FUNCTION CASENATP(Nat : String3) : Byte ;
begin { Pour les pièces }
Result:=7 ;
if Nat='FC' then Result:=1 else if Nat='AC' then Result:=2 else
 if ((Nat='RC') or (Nat='OC')) then Result:=3 else
  if Nat='FF' then Result:=4 else if Nat='AF' then Result:=5 else
   if ((Nat='RF') or (Nat='OF')) then Result:=6 else if Nat='ECC' then Result:=8 ;
end ;

{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
procedure ChangeAffGrid (GS : THGrid ; Mode : tmSais ; Decim : byte) ;
Var OBM : TOBM ;
    D,C : Double ;
    i   : integer ;
BEGIN
for i:=1 to GS.RowCount-2 do
    BEGIN
    OBM:=TOBM(GS.Objects[SA_Exo,i]) ; if OBM=Nil then Continue ;
    if Mode=tmPivot then
       BEGIN
       D:=OBM.GetMvt('E_DEBIT') ; C:=OBM.GetMvt('E_CREDIT') ;
       END else
       BEGIN
       D:=OBM.GetMvt('E_DEBITDEV') ; C:=OBM.GetMvt('E_CREDITDEV') ;
       END ;
    if D<>0 then
      begin
      GS.Cells[SA_Debit,i]:=StrfMontant(D,15,Decim,'',TRUE) ;
      GS.Cells[SA_Credit,i]:='';
      end
    else
      begin
      if Arrondi( C, Decim ) = 0
        then GS.Cells[SA_Credit,i]:=''
        else GS.Cells[SA_Credit,i]:=StrfMontant(C,15,Decim,'',TRUE) ;
      GS.Cells[SA_Debit,i]:='';
      end ;
    END ;
END ;

Procedure InitDefautColSaisie ;
BEGIN
SA_Exo      :=0  ; SA_NumP     :=1  ; SA_DateC    :=2  ; SA_NumL     :=3 ;
SA_Gen      :=4  ; SA_Aux      :=5  ; SA_RefI     :=6  ; SA_Lib      :=7 ;
SA_Debit    :=8  ; SA_Credit   :=9  ; SA_ModeP    :=10 ; SA_DateEche :=11 ;
SA_RIB      :=12 ; SA_RefE     :=13 ; SA_RefL     :=14 ; SA_NumTraChq:=15 ;
END ;

Procedure MajRib(Cpt,Etab,Guichet,Numero,Cle,Dom : String) ;
Var Principal,StSQL : String ;
    LastNum : Integer ;
    QRib : TQuery ;
    InsertAFaire : Boolean ;
    RibTrouve,UpdateFait : Boolean ;
BEGIN
If Trim(Cpt)='' Then Exit ;
Etab:=Trim(Etab) ; Guichet:=Trim(Guichet) ; Numero:=Trim(Numero) ; Cle:=Trim(Cle) ;
Dom:=Trim(Dom) ;
If (Etab='') Or (Guichet='') Or (Numero='') Then Exit ;
StSQL:='Select * from RIB Where R_AUXILIAIRE="'+Cpt+'"' ;
QRib:=OpenSQL(StSQL,FALSE) ;
InsertAFaire:=TRUE ; LastNum:=0 ; Principal:='X' ; UpdateFait:=FALSE ;
While Not QRib.Eof Do
  BEGIN
  RibTrouve:=FALSE ;
  If (QRib.Findfield('R_ETABBQ').ASString=Etab) And
     (QRib.Findfield('R_GUICHET').ASString=Guichet) And
     (QRib.Findfield('R_NUMEROCOMPTE').ASString=Numero) Then BEGIN InsertAFaire:=FALSE ; RibTrouve:=TRUE ; END ;
  If QRib.Findfield('R_PRINCIPAL').ASString='X' Then Principal:='-' ;
  If QRib.Findfield('R_NUMERORIB').ASInteger>LastNum Then LastNum:=QRib.Findfield('R_NUMERORIB').ASInteger ;
  If RibTrouve And (Not UpdateFait) Then //(QRib.Findfield('R_DOMICILIATION').ASString=Dom) And (QRib.Findfield('R_CLERIB').ASString=Cle)
    BEGIN
    QRib.Edit ;
    QRib.Findfield('R_CLERIB').ASString:=Cle ;
    If (Dom='') Then Dom:=QRib.Findfield('R_DOMICILIATION').ASString ;
    If Dom<>QRib.Findfield('R_DOMICILIATION').ASString Then QRib.Findfield('R_DOMICILIATION').ASString:=Dom ;
    QRib.Post ;
    UpdateFait:=TRUE ;
    END ;
  QRib.Next ;
  END ;
If InsertAFaire Then
   BEGIN
   QRib.insert ;
   InitNew(QRib) ;
   If Dom='' Then Dom:='A REMPLIR' ;
   QRib.Findfield('R_AUXILIAIRE').ASString:=Cpt ;
   QRib.Findfield('R_NUMERORIB').ASInteger:=LastNum+1 ;
   QRib.Findfield('R_DOMICILIATION').ASString:=Dom ;
   QRib.Findfield('R_ETABBQ').ASString:=Etab ;
   QRib.Findfield('R_GUICHET').ASString:=Guichet ;
   QRib.Findfield('R_NUMEROCOMPTE').ASString:=Numero ;
   QRib.Findfield('R_CLERIB').ASString:=Cle ;
   QRib.Findfield('R_PRINCIPAL').ASString:=Principal ;
   QRib.Findfield('R_VILLE').ASString:='A REMPLIR' ;
   QRib.Findfield('R_DEVISE').ASString:=V_PGI.DevisePivot ;
   QRib.Findfield('R_PAYS').ASString:='' ;
   QRib.Post ;
   END ;
Ferme(QRib) ;
END ;
{$ENDIF}
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : GP
Créé le ...... : 28/07/2000
Modifié le ... :   /  /
Description .. : Permet de controler un rib saisi avec ou sans séparateur.
Suite ........ : La fonction renvoie le code erreur sous forme integer.
Suite ........ : Si ok, result=0
Suite ........ : Function appelée notamment dans la saisie des effets en
Suite ........ : retour d'acceptation
Suite ........ : Si demande de création dans la table RIB, alors transaction
Suite ........ : sur la sous procedure MajRibTiers avec Création ou Modif
Suite ........ : (modif dans le cas ou le rib existe dans la table,
Suite ........ : mais avec une clé fausse)
Mots clefs ... :
*****************************************************************}
{$IFNDEF SANSCOMPTA}
{$IFNDEF PGIIMMO}
Function  FormatRIBEcr(Cpt : String ; AvecMess,VerifDansTable : Boolean ; Var St : String) : Integer ;
Const StErr : Array[1..6] Of String =
             ('Code banque non renseignée',
              'Longueur code banque erroné',
              'Code guichet non renseignée',
              'Longueur code Guichet erroné',
              'Code compte non renseignée',
              'Longueur code compte erroné'
              ) ;
Var St1,St2,St3 : String ;
    OkSep,NoSep,Pb : Boolean ;
    i,NbT : Integer ;
    Etab,Guichet,Numero,Cle,Dom,VraiCle : String ;
    SetErr : Set Of Byte ;
BEGIN // FormatRibEcr
Result:=0 ;
If ((Trim(St)='') or (st[1]='*')) Then Exit ;  // VL 06/05/2003 Pour l'iban
NoSep:=TRUE ; Etab:='' ; Guichet:='' ; Numero:='' ; Cle:='' ; Dom:='' ;
For i:=1 To length(St) Do
  BEGIN
  OkSep:=St[i] in ['/','-','.',',',';','=','+'] ;
  If OkSep Then BEGIN St[i]:=';' ; NoSep:=FALSE ; END ;
  END ;
If NoSep Then
  BEGIN
  Etab:=Copy(St,1,5) ; Guichet:=Copy(St,6,5) ; Numero:=Copy(St,11,11) ;
  Cle:=Copy(St,22,2) ; Dom:=Copy(St,28,24) ;
  END Else
  BEGIN
  St1:=St ; i:=1 ; If St1[Length(St1)]<>';' Then St1:=St1+';' ;
  While St1<>'' Do
    BEGIN
    St2:=ReadTokenSt(St1) ;
    Case i of 1 : Etab:=St2 ; 2 : Guichet:=St2 ; 3 : Numero:=St2 ; 4 : Cle:=St2 ; 5 : Dom:=St2 ; END ;
    Inc(i) ;
    END ;
  END ;
SetErr:=[] ;
If Etab='' Then BEGIN SetErr:=SetErr+[1] ; Result:=1 ; END ;
If Length(Etab)<>5 Then BEGIN SetErr:=SetErr+[2] ; Result:=2 ; END ;
If Guichet='' Then BEGIN SetErr:=SetErr+[3] ; Result:=3 ; END ;
If Length(Guichet)<>5 Then BEGIN SetErr:=SetErr+[4] ; Result:=4 ; END ;
If Numero='' Then BEGIN SetErr:=SetErr+[5] ; Result:=5 ; END ;
If Length(Numero)<>11 Then BEGIN SetErr:=SetErr+[6] ; Result:=6 ; END ;
//If Cle='' Then Result:=7 ;
//If Length(Cle)<>2 Then Result:=8 ;
Repeat
  If SetErr<>[] Then
    BEGIN
    For i:=1 To 6 Do If Byte(i) in SetErr Then Result:=i ;
    If Result=0 Then Break ;
    If (Byte(1) in SetErr) Or (Byte(3) in SetErr) Or (Byte(5) in SetErr) Then
      BEGIN
      HShowMessage('0;Attention : RIB incorrect;'+Sterr[Result]+';E;O;O;O;','','') ; Exit ;
      END Else
      BEGIN
      If HShowMessage('0;Attention : '+StErr[Result]+';Confirmez-vous ?;Q;YN;Y;N;O;','','')=6(*mryes*) then
        BEGIN
        SetErr:=SetErr-[Result] ;
        END Else
        BEGIN
        HShowMessage('0;Attention : RIB incorrect;'+Sterr[Result]+';E;O;O;O;','','') ; Exit ;
        END ;
      END ;
    END ;
Until SetErr=[] ;
Result:=0 ;
VraiCle:=VerifRib(Etab,Guichet,Numero) ;
If (VraiCle<>Cle)  Then
  BEGIN
  St3:='0;Attention : Clé RIB incorrecte;Voulez-vous la Changer ?( '+VraiCle+' Au lieu de '+Cle+' );Q;YN;Y;N;O;' ;
  If HShowMessage(St3,'','')=6(*mryes*) then Cle:=VraiCle ;
  END ;
If VerifDansTable Then
  BEGIN
  NbT:=1 ;
  Repeat
   Try
    Pb:=FALSE ; Inc(NbT) ;
    BeginTrans ;
    MajRib(Cpt,Etab,Guichet,Numero,Cle,Dom) ;
    CommitTrans ;
   Except
    RollBack ; Pb:=TRUE ;
   End ;
  Until (Not Pb) Or (NbT>5) ;
  END ;
If Result=0 Then St:=EncodeRib(Etab,Guichet,Numero,Cle,Dom) ;
END ; // FormatRibEcr
{$ENDIF}
{$ENDIF}

{$IFNDEF NOVH}
Procedure SauveSAEff ;
BEGIN
VH^.SauveSA_SaisieEff[0]:=SA_Exo   ;
VH^.SauveSA_SaisieEff[1]:=SA_NumP  ;
VH^.SauveSA_SaisieEff[2]:=SA_DateC ;
VH^.SauveSA_SaisieEff[3]:=SA_NumL  ;
VH^.SauveSA_SaisieEff[4]:=SA_Gen   ;
VH^.SauveSA_SaisieEff[5]:=SA_Aux   ;
VH^.SauveSA_SaisieEff[6]:=SA_RefI  ;
VH^.SauveSA_SaisieEff[7]:=SA_Lib   ;
VH^.SauveSA_SaisieEff[8]:=SA_Debit ;
VH^.SauveSA_SaisieEff[9]:=SA_Credit;
VH^.SauveSA_SaisieEff[10]:=SA_ModeP ;
VH^.SauveSA_SaisieEff[11]:=SA_DateEche ;
VH^.SauveSA_SaisieEff[12]:=SA_RIB      ;
VH^.SauveSA_SaisieEff[13]:=SA_RefE     ;
VH^.SauveSA_SaisieEff[14]:=SA_RefL     ;
VH^.SauveSA_SaisieEff[15]:=SA_NumTraChq;
END ;

Procedure RestoreSAEff ;
BEGIN
SA_Exo:=VH^.SauveSA_SaisieEff[0]   ;
SA_NumP:=VH^.SauveSA_SaisieEff[1]  ;
SA_DateC:=VH^.SauveSA_SaisieEff[2] ;
SA_NumL:=VH^.SauveSA_SaisieEff[3]  ;
SA_Gen:=VH^.SauveSA_SaisieEff[4]   ;
SA_Aux:=VH^.SauveSA_SaisieEff[5]   ;
SA_RefI:=VH^.SauveSA_SaisieEff[6]  ;
SA_Lib:=VH^.SauveSA_SaisieEff[7]   ;
SA_Debit:=VH^.SauveSA_SaisieEff[8] ;
SA_Credit:=VH^.SauveSA_SaisieEff[9];
SA_ModeP:=VH^.SauveSA_SaisieEff[10] ;
SA_DateEche:=VH^.SauveSA_SaisieEff[11] ;
SA_RIB:=VH^.SauveSA_SaisieEff[12]      ;
SA_RefE:=VH^.SauveSA_SaisieEff[13]     ;
SA_RefL:=VH^.SauveSA_SaisieEff[14]     ;
SA_NumTraChq:=VH^.SauveSA_SaisieEff[15] ;
END ;

Procedure initSA_SaisieNormale ;
BEGIN
SA_Exo      := 0 ;
SA_NumP     := 1 ;
SA_DateC    := 2 ;
SA_NumL     := 3 ;
SA_Gen      := 4 ;
SA_Aux      := 5 ;
SA_RefI     := 6 ;
SA_Lib      := 7 ;
SA_Debit    := 8 ;
SA_Credit   := 9 ;
SA_ModeP    := 10 ;
SA_DateEche := 11 ;
SA_RIB      := 12 ;
SA_RefE     := 13 ;
SA_RefL     := 14 ;
SA_NumTraChq:= 15 ;
END ;
{$ENDIF}

{$IFNDEF COMSX}
{***********A.G.L.***********************************************
Auteur  ...... : Xavier Maluenda
Créé le ...... : 24/02/2005
Modifié le ... :   /  /
Description .. : Si pays de localisation <>ESP, on renvoie TOUJOURS
Suite ........ : TRUE, si Espagne, on "regarde" le tiers pour repondre....
Suite ........ :
Suite ........ :
Suite ........ : HISTORIQUE DE MODIFICATIONS:
Suite ........ : 24/02/2005 (XVI) re-unification des sourcecs ESP et FRA
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
Function IsTiersSoumisTPF ( pGS : THGrid ) : Boolean ;
var IdxAux : Integer ;
Begin
  if (VH^.PaysLocalisation<>CodeISOES) then
     result:=TRUE
  else
  begin
    Result:=FALSE ;
    idxAux:=TrouveLigTiers(pGS,0) ;
    if idxAux>0 then
       Result:=GetGTiers(pGS,idxAux).SoumisTPF ;
  End ;
End ;
{$ENDIF}
{$ENDIF}


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 24/05/2005
Modifié le ... :   /  /
Description .. : Uniquement en mode MULTISOC !
Suite ........ :
Suite ........ : Maj des totaux des comptes dans les tables locales
Mots clefs ... :
*****************************************************************}
(* ANCIENNE VERSION
function CUpdateCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vTotDeb, vTotCre, vTotDP, vTotCP, vTotDE, vTotCE, vTotDS, vTotCS : Double ; vBoAno : Boolean ; vDossier : String = '' ) : Boolean ;
var lTobCumul : TOB ;
    vBoInsert : Boolean ;
    lQCumul   : TQuery ;
begin

  result := True ;

  if not EstTablePartagee( fbToTable( vFB ) ) then Exit ;

  if vFB <> fbSect then vAxe := '' ; // Certains traitements utilise un identifiant T / G ... dans l'axe...

  // =====================================
  // === Recherche de l'enregistrement ===
  // =====================================

  lTobCumul := Tob.Create( 'CUMULS', nil, -1 ) ;
  if (vDossier = '') or (vDossier = V_PGI.SchemaName) then
    begin
    lTobCumul.PutValue('CU_TYPE',      fbToCumulType( vFB ) ) ;
    lTobCumul.PutValue('CU_COMPTE1',   vCompte ) ;
    lTobCumul.PutValue('CU_COMPTE2',   vAxe ) ;
    vBoInsert := not lTobCumul.LoadDB ;
    end
  else
    begin
    lQCumul := OpenSelect( 'SELECT * FROM CUMULS WHERE CU_TYPE="'    + fbToCumulType( vFB ) + '"'
                                               + ' AND CU_COMPTE1="' + vCompte + '"'
                                               + ' AND CU_COMPTE2="' + vAxe + '"', vDossier ) ;
    vBoInsert := not lTobCumul.SelectDB('', lQCumul ) ;
    Ferme(lQCumul) ;
    end ;

  // ======================
  // === MAJ des totaux ===
  // ======================

  // _TOTDEBIT / _TOTALCREDIT
  lTobCumul.SetDouble('CU_DEBIT1',  lTobCumul.GetDouble('CU_DEBIT1') + vTotDeb ) ;
  lTobCumul.SetDouble('CU_CREDIT1', lTobCumul.GetDouble('CU_CREDIT1') + vTotCre ) ;

  if vBoAno then
    begin
    // _TOTDEBAN / _TOTCREAN
    lTobCumul.SetDouble('CU_DEBITAN',  lTobCumul.GetDouble('CU_DEBITAN') + vTotDeb ) ;
    lTobCumul.SetDouble('CU_CREDITAN', lTobCumul.GetDouble('CU_CREDITAN') + vTotCre ) ;
    // _TOTDEBE / _TOTCREE
    lTobCumul.SetDouble('CU_DEBIT3',    lTobCumul.GetDouble('CU_DEBIT3') + vTotDeb ) ;
    lTobCumul.SetDouble('CU_CREDIT3',   lTobCumul.GetDouble('CU_CREDIT3') + vTotCre ) ;
    end
  else
    begin
    // _TOTDEBP / _TOTCREP
    lTobCumul.SetDouble('CU_DEBIT2',  lTobCumul.GetDouble('CU_DEBIT2') + vTotDP ) ;
    lTobCumul.SetDouble('CU_CREDIT2', lTobCumul.GetDouble('CU_CREDIT2') + vTotCP ) ;
    // _TOTDEBE / _TOTCREE
    lTobCumul.SetDouble('CU_DEBIT3',  lTobCumul.GetDouble('CU_DEBIT3') + vTotDE ) ;
    lTobCumul.SetDouble('CU_CREDIT3', lTobCumul.GetDouble('CU_CREDIT3') + vTotCE ) ;
    // _TOTDEBS / _TOTCRES
    lTobCumul.SetDouble('CU_DEBIT4',  lTobCumul.GetDouble('CU_DEBIT4') + vTotDS ) ;
    lTobCumul.SetDouble('CU_CREDIT4', lTobCumul.GetDouble('CU_CREDIT4') + vTotCS ) ;
    end ;

  if vBoInsert then
    begin
    lTobCumul.PutValue('CU_QUALIFPIECE', 'N' ) ;
    lTobCumul.PutValue('CU_DEVQTE',      V_PGI.DevisePivot ) ;
    end ;

  // ================
  // === MAJ BASE ===
  // ================
  if (vDossier = '') or (vDossier = V_PGI.SchemaName) then
    begin
    if vBoInsert
      then result := lTobCumul.InsertDB(nil)
      else result := lTobCumul.UpdateDB ;
    end
  else
    begin
    if vBoInsert
      then result := InsertTobMS( lTobCumul, vDossier )
      else result := UpdateTobMS( lTobCumul, vDossier ) ;
    end ;

  if not result then
    V_PGI.IoError:=oeSaisie;

  FreeAndNil( lTobCumul ) ;

end ; *)

function CUpdateCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vTotDeb, vTotCre, vTotDP, vTotCP, vTotDE, vTotCE, vTotDS, vTotCS : Double ; vBoAno : Boolean ; vDossier : String = '' ) : Boolean ;
var lStReq    : string ;
    lStWhere  : string ;
begin

  result := True ;

  // Traitement spécial si appel pour un axe
  if vFb in [fbAxe1..fbAXe5] then
    begin
    vAxe := fbToAxe( vFb ) ;
    vFb  := fbSect ;
    end
  else if vFb <> fbSect then vAxe := '' ;// Certains traitements utilise un identifiant T / G ... dans l'axe...

  if not EstTablePartagee( fbToTable( vFB ) ) then Exit ;

  lStWhere :=  ' WHERE CU_TYPE="' + fbToCumulType( vFB ) + '"'
               + ' AND CU_COMPTE1="' + vCompte + '"'
               + ' AND CU_COMPTE2="' + vAxe + '"' ;

  // ====================================================
  // === Test présence du cumul et création si besoin ===
  // ====================================================
  if not CCreateCumulsMS( vFB, vCompte, vAxe, vDossier ) then
    begin
    result        := False ;
    V_PGI.IoError := oeSaisie;
    Exit ;
    end ;

  // ==========================================
  // === Création requête de MAJ des totaux ===
  // ==========================================

  lStReq := 'UPDATE ' + GetTableDossier( vDossier, 'CUMULS' ) + ' SET ' ;
  // _TOTDEBIT / _TOTALCREDIT
  lStReq := lStReq   + 'CU_DEBIT1 = CU_DEBIT1 + ' + StrFPoint( vTotDeb )
                   + ', CU_CREDIT1 = CU_CREDIT1 + ' + StrFPoint( vTotCre ) ;

  if vBoAno then
    begin
    // _TOTDEBAN / _TOTCREAN
    lStReq := lStReq + ', CU_DEBITAN = CU_DEBITAN + ' + StrFPoint( vTotDeb )
                     + ', CU_CREDITAN = CU_CREDITAN + ' + StrFPoint( vTotCre ) ;
    // _TOTDEBE / _TOTCREE
    lStReq := lStReq + ', CU_DEBIT3 = CU_DEBIT3 + ' + StrFPoint( vTotDeb )
                     + ', CU_CREDIT3 = CU_CREDIT3 + ' + StrFPoint( vTotCre ) ;
    end
  else
    begin
    // _TOTDEBP / _TOTCREP
    lStReq := lStReq + ', CU_DEBIT2 = CU_DEBIT2 + ' + StrFPoint( vTotDP )
                     + ', CU_CREDIT2 = CU_CREDIT2 + ' + StrFPoint( vTotCP ) ;
    // _TOTDEBE / _TOTCREE
    lStReq := lStReq + ', CU_DEBIT3 = CU_DEBIT3 + ' + StrFPoint( vTotDE )
                     + ', CU_CREDIT3 = CU_CREDIT3 + ' + StrFPoint( vTotCE ) ;
    // _TOTDEBS / _TOTCRES
    lStReq := lStReq + ', CU_DEBIT4 = CU_DEBIT4 + ' + StrFPoint( vTotDS )
                     + ', CU_CREDIT4 = CU_CREDIT4 + ' + StrFPoint( vTotCS ) ;
    end ;

  lStReq := lStReq + lStWhere ;

  // ================
  // === MAJ BASE ===
  // ================
  result := ExecuteSQL( lStReq ) = 1 ;

  if not result then
    V_PGI.IoError:=oeSaisie;

end ;

procedure CChargeCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; var vTobCpt : TOB ; vDossier : String = '' ) ;
var lTobTot   : TOB ;
    lNomTable : String ;
    lPref     : String ;
    lBoLoad   : Boolean ;
    lQCumul   : TQuery ;
begin

  Case vFB of
    fbGene : begin
             lNomTable := 'GENERAUX' ;
             lPref     := 'G' ;
             end ;
    fbAux  : begin
             lNomTable := 'TIERS' ;
             lPref     := 'T' ;
             end ;
    fbSect : begin
              if VH^.LiaisonY2ViaShare then
              begin
               lNomTable := 'CSECTION' ;
               lPref     := 'CSP' ;
              end else
              begin
               lNomTable := 'SECTION' ;
               lPref     := 'S' ;
              end;
             end ;
    else Exit ;
    end ;

  if not EstTablePartagee( lNomTable ) then Exit ;

  lTobTot := Tob.Create('CUMULS', nil, -1 ) ;
  if (vDossier = '') or (vDossier = V_PGI.SchemaName) then
    begin
    lTobTot.SetString('CU_TYPE',         fbToCumulType( vFB ) ) ;
    lTobTot.SetString('CU_COMPTE1',      vCompte ) ;
    lTobTot.SetString('CU_COMPTE2',      vAxe ) ;
    lBoLoad := lTobTot.LoadDB ;
    end
  else
    begin
    lQCumul := OpenSelect( 'SELECT * FROM CUMULS WHERE CU_TYPE="'    + fbToCumulType( vFB ) + '"'
                                               + ' AND CU_COMPTE1="' + vCompte + '"'
                                               + ' AND CU_COMPTE2="' + vAxe + '"', vDossier ) ;
    lBoLoad := lTobTot.SelectDB('', lQCumul ) ;
    Ferme(lQCumul) ;
    end ;


  if lBoLoad then
    begin
    // TOTALDEBIT
    vTobCpt.SetDouble( lPref + '_TOTALDEBIT',    lTobTot.GetDouble('CU_DEBIT1') ) ;
    vTobCpt.SetDouble( lPref + '_TOTALCREDIT',   lTobTot.GetDouble('CU_CREDIT1') ) ;
    // TOTP
    vTobCpt.SetDouble( lPref + '_TOTDEBP',       lTobTot.GetDouble('CU_DEBIT2') ) ;
    vTobCpt.SetDouble( lPref + '_TOTCREP',       lTobTot.GetDouble('CU_CREDIT2') ) ;
    // TOTE
    vTobCpt.SetDouble( lPref + '_TOTDEBE',       lTobTot.GetDouble('CU_DEBIT3') ) ;
    vTobCpt.SetDouble( lPref + '_TOTCREE',       lTobTot.GetDouble('CU_CREDIT3') ) ;
    // TOTS
    vTobCpt.SetDouble( lPref + '_TOTDEBS',       lTobTot.GetDouble('CU_DEBIT4') ) ;
    vTobCpt.SetDouble( lPref + '_TOTCRES',       lTobTot.GetDouble('CU_CREDIT4') ) ;

    if vFB <> fbJal then
      begin
      // TOTANO
      vTobCpt.SetDouble( lPref + '_TOTDEBANO',     lTobTot.GetDouble('CU_DEBITAN') ) ;
      vTobCpt.SetDouble( lPref + '_TOTCREANO',     lTobTot.GetDouble('CU_CREDITAN') ) ;
      // TOTANON1
      vTobCpt.SetDouble( lPref + '_TOTDEBANON1',   lTobTot.GetDouble('CU_DEBIT5') ) ;
      vTobCpt.SetDouble( lPref + '_TOTCREANON1',   lTobTot.GetDouble('CU_CREDIT5') ) ;
      end ;

    if vFB = fbGene then
      begin
      // NONPOINTE
      vTobCpt.SetDouble('G_DEBNONPOINTE',  lTobTot.GetDouble('CU_DEBIT6') ) ;
      vTobCpt.SetDouble('G_CREDNONPOINTE', lTobTot.GetDouble('CU_CREDIT6') ) ;
      // TOTPTP
      vTobCpt.SetDouble('G_TOTDEBPTP',     lTobTot.GetDouble('CU_DEBIT7') ) ;
      vTobCpt.SetDouble('G_TOTCREPTP',     lTobTot.GetDouble('CU_CREDIT7') ) ;
      // TOTPTD
      vTobCpt.SetDouble('G_TOTDEBPTD',     lTobTot.GetDouble('CU_DEBIT8') ) ;
      vTobCpt.SetDouble('G_TOTCREPTD',     lTobTot.GetDouble('CU_CREDIT8') ) ;
      end ;

    end
  else // Mise à zéro des montants si aucun enregistrement
    begin
    // TOTALDEBIT
    vTobCpt.SetDouble( lPref + '_TOTALDEBIT',    0 ) ;
    vTobCpt.SetDouble( lPref + '_TOTALCREDIT',   0 ) ;
    // TOTP
    vTobCpt.SetDouble( lPref + '_TOTDEBP',       0 ) ;
    vTobCpt.SetDouble( lPref + '_TOTCREP',       0 ) ;
    // TOTE
    vTobCpt.SetDouble( lPref + '_TOTDEBE',       0 ) ;
    vTobCpt.SetDouble( lPref + '_TOTCREE',       0 ) ;
    // TOTS
    vTobCpt.SetDouble( lPref + '_TOTDEBS',       0 ) ;
    vTobCpt.SetDouble( lPref + '_TOTCRES',       0 ) ;
    // TOTANO
    vTobCpt.SetDouble( lPref + '_TOTDEBANO',     0 ) ;
    vTobCpt.SetDouble( lPref + '_TOTCREANO',     0 ) ;
    // TOTANON1
    vTobCpt.SetDouble( lPref + '_TOTDEBANON1',   0 ) ;
    vTobCpt.SetDouble( lPref + '_TOTCREANON1',   0) ;

    if vFB = fbGene then
      begin
      // NONPOINTE
      vTobCpt.SetDouble('G_DEBNONPOINTE',  0) ;
      vTobCpt.SetDouble('G_CREDNONPOINTE', 0) ;
      // TOTPTP
      vTobCpt.SetDouble('G_TOTDEBPTP',     0) ;
      vTobCpt.SetDouble('G_TOTCREPTP',     0) ;
      // TOTPTD
      vTobCpt.SetDouble('G_TOTDEBPTD',     0) ;
      vTobCpt.SetDouble('G_TOTCREPTD',     0) ;
      // TOTDEBN2 / TOTCREN2
      vTobCpt.SetDouble('G_TOTDEBN2',     0) ;
      vTobCpt.SetDouble('G_TOTCREN2',     0) ;
      end ;

    end ;

  FreeAndNil( lTobTot ) ;

end ;

procedure CInsertCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vBoAll : Boolean = False) ;
var lTobLocal : Tob ;
    lStSQL    : String ;
begin

  if not EstTablePartagee( fbToTable( vFB ) ) then Exit ;

  if vFb <> fbSect then vAxe := '' ;

  // Ajout de l'enregistrement local de la table cumuls pour gestion des totaux en multisoc
  if not vBoAll then
    begin
    lTobLocal := Tob.Create('CUMULS', nil, -1);
    lTobLocal.PutValue('CU_TYPE',        fbToCumulType( vFB ) );
    lTobLocal.PutValue('CU_COMPTE1',     vCompte );
    lTobLocal.PutValue('CU_COMPTE2',     vAxe );
    lTobLocal.PutValue('CU_QUALIFPIECE', 'N' );
    lTobLocal.PutValue('CU_DEVQTE',      V_PGI.DevisePivot );
    lTobLocal.InsertDB(nil) ;
    FreeAndNil( lTobLocal ) ;
    end
  else
    // Cas spécifique du traitement de masse pour le recalcul des soldes des comptes
    begin
    // --> Requête <---
    lStSQL := 'INSERT INTO CUMULS SELECT "' + fbToCumulType( vFB ) + '", ' ;
    case vFB of
      fbGene : lStSQL := lStSQL + 'G_GENERAL, "", '  ;
      fbAux  : lStSQL := lStSQL + 'T_AUXILIAIRE, "", '  ;
      fbSect : lStSQL := lStSQL + 'S_SECTION, S_AXE, '  ;
      fbJal  : lStSQL := lStSQL + 'J_JOURNAL, "", '  ;
      end ;
    lStSQL := lStSQL + '"", "", "", "' + V_PGI.DevisePivot + '", "N", '
                     + '0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '
                     + '0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ""'
                     + ' FROM ' + fbToTable( vFB )
                     + ' WHERE NOT EXISTS( SELECT CU_COMPTE1 FROM CUMULS WHERE ' ;
    case vFB of
      fbGene : lStSQL := lStSQL + 'CU_COMPTE1 = G_GENERAL ) '  ;
      fbAux  : lStSQL := lStSQL + 'CU_COMPTE1 = T_AUXILIAIRE ) '  ;
      fbSect : lStSQL := lStSQL + 'CU_COMPTE1 = S_SECTION AND CU_COMPTE2 = S_AXE ) AND S_AXE = "' + vAxe + '"'  ;
      fbJal  : lStSQL := lStSQL + 'CU_COMPTE1 = J_JOURNAL ) '  ;
      end ;

    // --> MAJ Base <---
    ExecuteSQL( lStSQL ) ;

    end ;


end ;

procedure CDeleteCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vBoAll : Boolean ) ;
var lStSQL : String ;
begin
  // Uniquement pour les tables partagee en multisoc !
  if not EstTablePartagee( fbToTable( vFB ) ) then Exit ;

  lStSQL := 'DELETE FROM CUMULS WHERE CU_TYPE="' + fbToCumulType( vFB ) + '"' ;

  if not vBoAll then
    begin
    lStSQL := lStSQL + ' AND CU_COMPTE1="' + vCompte +'"' ;
    if vFB = fbSect
      then lStSQL := lStSQL + ' AND CU_COMPTE2="' + vAxe +'"' ;
    end ;

  ExecuteSQL( lStSQL ) ;

end ;


procedure CReinitCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vStCodeExo : String ; vBoODAnal : Boolean ; vBoAll : Boolean ) ;
var lStSQL    : String ;
    ExoN2     : TExoDate ;
begin

  // Traitement spécial si appel pour un axe
  if vFb in [fbAxe1..fbAXe5] then
    begin
    vAxe := fbToAxe( vFb ) ;
    vFb  := fbSect ;
    end
  else if vFb <> fbSect then vAxe := '' ;

  // Uniquement pour les tables partagee en multisoc !
  if not EstTablePartagee( fbToTable( vFB ) ) then Exit ;

{$IFNDEF NOVH}
  ExoN2 := CGetExerciceNMoins2;
{$ELSE}
  ExoN2 := GetNMoins2 ;
{$ENDIF NOVH}

    lStSQL := 'UPDATE CUMULS SET ' ;

    // Champs MAJ
    if (vStCodeExo <> GetPrecedent.Code) and (vStCodeExo <> ExoN2.Code) then
      lStSQL := lStSQL + 'CU_DEBIT1=0, CU_CREDIT1=0,';    // TOTALDEBIT / TOTALCREDIT

    if vStCodeExo = '' then
      begin
      lStSQL := lStSQL + 'CU_DEBIT2=0, CU_CREDIT2=0,'             // TOTDEBP / TOTCREP
                       + 'CU_DEBIT3=0, CU_CREDIT3=0,'             // TOTDEBE / TOTCREE
                       + 'CU_DEBIT4=0, CU_CREDIT4=0';            // TOTDEBS / TOTCRES
      if vFB<>fbJal then
        begin
        lStSQL := lStSQL + ', CU_DEBITAN=0, CU_CREDITAN=0'         // TOTDEBANO , TOTCREANO
                         + ', CU_DEBIT5=0, CU_CREDIT5=0' ;         // TOTDEBANON1 , TOTCREANON1
        if vFB=fbGene then
          lStSQL := lStSQL + ', CU_DEBIT9=0, CU_CREDIT9=0'         // TOTDEBN2 , TOTCREN2
        end ;
      end
    else
      begin
      if vStCodeExo = GetPrecedent.Code then
        lStSQL := lStSQL + 'CU_DEBIT2=0, CU_CREDIT2=0'
      else if vStCodeExo = GetEnCours.Code   then
        begin
        lStSQL := lStSQL + 'CU_DEBIT3=0, CU_CREDIT3=0' ;
        if vFb <> fbJal then
          lStSQL := lStSQL + ', CU_DEBITAN=0, CU_CREDITAN=0'
        end
      else if vStCodeExo = GetSuivant.Code   then
        begin
        lStSQL := lStSQL + 'CU_DEBIT4=0, CU_CREDIT4=0' ;
        if vFb <> fbJal then
          lStSQL := lStSQL + ', CU_DEBIT5=0, CU_CREDIT5=0'
        end
      else if vStCodeExo = ExoN2.Code  then
        lStSQL := lStSQL + 'CU_DEBIT9=0, CU_CREDIT9=0' ;
      end ;


    // ====================
    // ==== Conditions ====
    // ====================
    // type de cumuls
    lStSQL := lStSQL + ' WHERE CU_TYPE="' + fbToCumulType( vFB ) + '"' ;
    // Axe pour les sections
    if vFB = fbSect
      then lStSQL := lStSQL + ' AND CU_COMPTE2="' + vAxe +'"' ;
    // pour un compte précis?
    if not vBoAll then
      begin
      lStSQL := lStSQL + ' AND CU_COMPTE1="' + vCompte +'"' ;
      end
    else if (vFB = fbJal) then
      // Distinction des journaux d'ODA des journaux classiques
      if vBoODAnal
        then lStSQL := lStSQL + ' AND CU_COMPTE1 IN ( SELECT J_JOURNAL FROM JOURNAL WHERE (J_NATUREJAL="ODA" OR J_NATUREJAL="ANA") AND J_AXE="'+ vAxe +'")'
        else lStSQL := lStSQL + ' AND CU_COMPTE1 IN ( SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL<>"ODA" And J_NATUREJAL<>"ANA" AND J_AXE="'+ vAxe +'")' ;

  ExecuteSQL( lStSQL ) ;

  // Insertion des cumuls non existants
  if vBoAll then
    CInsertCumulsMS( vFB, '', vAxe, True ) ;

end ;


procedure CMajTotCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vTotDeb, vTotCre : Double ; vTabDeb, vTabCre : Array Of Double ; vBoInit : Boolean ) ;
Var DD,CD              : String;
    DP,CP              : String;
    DE,CE              : String;
    DS,CS              : String;
    DA,CA              : String;
    DA1,CA1            : String;
    DNMoins2, CNMoins2 : String;
    lStSQL             : String ;
    lStWhere           : String ;
begin

  // Traitement spécial si appel pour un axe
  if vFb in [fbAxe1..fbAxe5] then
    begin
    vAxe := fbToAxe( vFb ) ;
    vFb  := fbSect ;
    end
  else if vFb <> fbSect then vAxe := '' ;

  // Uniquement pour les tables partagee en multisoc !
  if not EstTablePartagee( fbToTable( vFB ) ) then Exit ;


  // ==== Conditions sur le cumul ====
  // type de cumuls
  lStWhere := lStWhere + ' WHERE CU_TYPE="' + fbToCumulType( vFB ) + '"' ;
  // Code du compte
  lStWhere := lStWhere + ' AND CU_COMPTE1="' + vCompte +'"' ;
  // Axe pour les sections
  if vFB = fbSect
    then lStWhere := lStWhere + ' AND CU_COMPTE2="' + vAxe +'"' ;

  // -------------------------------------------------------------------
  // --- Récupération des montants (cf FabricReqCpt de SoldeCpt.pas) ---
  // -------------------------------------------------------------------
  DD := VariantToSQL( vTotDeb );
  CD := VariantToSQL( vTotCre );
  DP := VariantToSQL( vTabDeb[0] );
  CP := VariantToSQL( vTabCre[0] );
  DE := VariantToSQL( vTabDeb[1] );
  CE := VariantToSQL( vTabCre[1] );
  DS := VariantToSQL( vTabDeb[2] );
  CS := VariantToSQL( vTabCre[2] );
  If vFB <> fbJal then
    begin
    DA  := VariantToSQL( vTabDeb[3] );
    CA  := VariantToSQL( vTabCre[3] );
    DA1 := VariantToSQL( vTabDeb[4] );
    CA1 := VariantToSQL( vTabCre[4] );
    end ;
  If vFB = fbGene then
    begin
    DNMoins2 := VariantToSQL( vTabDeb[5] );
    CNMoins2 := VariantToSQL( vTabCre[5] );
    end ;

  // ---------------
  // --- Requête ---
  // ---------------

    // --- Si le cumul n'existe pas, on le crée ---
    if not ExisteSQL('SELECT CU_COMPTE1 FROM CUMULS' + lStWhere ) then
      begin
      lStSQL := 'INSERT INTO CUMULS ( CU_TYPE, CU_COMPTE1, CU_COMPTE2, CU_EXERCICE, CU_SUITE, CU_ETABLISSEMENT, CU_DEVQTE, CU_QUALIFPIECE,'
                                  + ' CU_DEBIT1, CU_CREDIT1, CU_DEBIT2, CU_CREDIT2, CU_DEBIT3, CU_CREDIT3, CU_DEBIT4, CU_CREDIT4,'
                                  + ' CU_DEBIT5, CU_CREDIT5, CU_DEBIT6, CU_CREDIT6, CU_DEBIT7, CU_CREDIT7, CU_DEBIT8, CU_CREDIT8,'
                                  + ' CU_DEBIT9, CU_CREDIT9, CU_DEBIT10, CU_CREDIT10, CU_DEBIT11, CU_CREDIT11, CU_DEBIT12, CU_CREDIT12,'
                                  + ' CU_CREDITAN, CU_DEBITAN, CU_QUALIFQTEDIFF ) ' ;

      // Identifiants
      lStSQL := lStSQL + ' VALUES ( "' + fbToCumulType( vFB ) + '", "' + vCompte + '", "' + vAxe + '", "", "", "", "' + V_PGI.DevisePivot + '", "N", ' ;
      // Tronc commun : montants 1 à 4 (N,P,E,S)
      lStSQL := lStSQL + DD + ', ' + CP + ', ' + DP + ', ' + CP + ', ' + DE + ', ' + CE + ', ' + DS + ', ' + CS + ', ' ;
      // Montants 5-12 + AN suivant type de compte
      if vFB = fbJal then
        // tout le reste à 0 pour les journaux
        lStSQL := lStSQL + '0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '
      else
        begin
        // montant A-Nouveeaux ( + N-2 pour le généraux)
        lStSQL := lStSQL + DA1 + ', '+ CA1 + ', 0, 0, 0, 0, 0, 0, ' ;
        if vFB = fbGene
          then lStSQL := lStSQL + DNMoins2 + ', ' + CNMoins2 + ', '
          else lStSQL := lStSQL + '0, 0, ' ;
        lStSQL := lStSQL + '0, 0, 0, 0, 0, 0, ' + CA + ', ' + DA + ', ' ;
        end ;

      // dernier champs
      lStSQL := lStSQL + '"" )' ;
      end
    // Le cumul existe en base
    else if vBoInit then
      begin
        lStSQL := 'UPDATE CUMULS SET ' ;
        // ==== Champs à mettre à jour ====
        // TOTALDEBIT / TOTALCREDIT
        lStSQL := lStSQL + 'CU_DEBIT1=' + DD + ', CU_CREDIT1=' + CD ;
        // TOTDEBP / TOTCREP
        lStSQL := lStSQL + ', CU_DEBIT2=' + DP + ', CU_CREDIT2=' + CP ;
        // TOTDEBE / TOTCREE
        lStSQL := lStSQL + ', CU_DEBIT3=' + DE + ', CU_CREDIT3=' + CE ;
        // TOTDEBS / TOTCRES
        lStSQL := lStSQL + ', CU_DEBIT4=' + DS + ', CU_CREDIT4=' + CS ;
        if vFB <> fbJal then
          begin
          // TOTDEBANO , TOTCREANO
          lStSQL := lStSQL + ', CU_DEBITAN=' + DA + ', CU_CREDITAN=' + CA ;
          // TOTDEBANON1 , TOTCREANON1
          lStSQL := lStSQL + ', CU_DEBIT5=' + DA1 + ', CU_CREDIT5=' + CA1 ;
          end ;
        // TOTDEBN2 , TOTCREN2
        if vFb = fbGene then
          lStSQL := lStSQL + ', CU_DEBIT9=' + DNMoins2 + ', CU_CREDIT9=' + CNMoins2 ;

        // === AJOUT DE LA CONDITION
        lStSQL := lStSQL + lStWhere ;
      end
  else
  // --- Sinon on update ---
    begin
    lStSQL := 'UPDATE CUMULS SET ' ;
    // ==== Champs à mettre à jour ====
    // TOTALDEBIT / TOTALCREDIT
    lStSQL := lStSQL + 'CU_DEBIT1=CU_DEBIT1+' + DD + ', CU_CREDIT1=CU_CREDIT1+' + CD ;
    // TOTDEBP / TOTCREP
    lStSQL := lStSQL + ', CU_DEBIT2=CU_DEBIT2+' + DP + ', CU_CREDIT2=CU_CREDIT2+' + CP ;
    // TOTDEBE / TOTCREE
    lStSQL := lStSQL + ', CU_DEBIT3=CU_DEBIT3+' + DE + ', CU_CREDIT3=CU_CREDIT3+' + CE ;
    // TOTDEBS / TOTCRES
    lStSQL := lStSQL + ', CU_DEBIT4=CU_DEBIT4+' + DS + ', CU_CREDIT4=CU_CREDIT4+' + CS ;
    if vFB <> fbJal then
      begin
      // TOTDEBANO , TOTCREANO
      lStSQL := lStSQL + ', CU_DEBITAN=CU_DEBITAN+' + DA + ', CU_CREDITAN=CU_CREDITAN+' + CA ;
      // TOTDEBANON1 , TOTCREANON1
      lStSQL := lStSQL + ', CU_DEBIT5=CU_DEBIT5+' + DA1 + ', CU_CREDIT5=CU_CREDIT5+' + CA1 ;
      end ;
    // TOTDEBN2 , TOTCREN2
    if vFb = fbGene then
      lStSQL := lStSQL + ', CU_DEBIT9=CU_DEBIT9+' + DNMoins2 + ', CU_CREDIT9=CU_CREDIT9+' + CNMoins2 ;

    // === AJOUT DE LA CONDITION
    lStSQL := lStSQL + lStWhere ;
    end ;

  // ----------------
  // --- MAJ base ---
  // ----------------
  ExecuteSQL( lStSQL ) ;

end ;


function  CUpdateCumulsPointeMS( vCompte : String ; vTotDebPTP, vTotCrePTP, vTotDebPTD, vTotCrePTD : Double ; vBoInit : Boolean = True ; vDossier : String = '' ) : Boolean ;
var lStReq   : String ;
    lStWhere : String ;
begin

  result := True ;

  if not EstTablePartagee( 'GENERAUX' ) then Exit ;

  lStWhere :=  ' WHERE CU_TYPE="' + fbToCumulType( fbGene ) + '"'
               + ' AND CU_COMPTE1="' + vCompte + '"'
               + ' AND CU_COMPTE2=""' ;

  // ========================================
  // === Test présence du cumul si besoin ===
  // ========================================
  if not CCreateCumulsMS( fbGene, vCompte, '', vDossier ) then
    begin
    result        := False ;
    V_PGI.IoError := oeSaisie;
    Exit ;
    end ;


  // =================================
  // === Requête de MAJ des totaux ===
  // =================================

  lStReq := 'UPDATE ' + GetTableDossier( vDossier, 'CUMULS' ) + ' SET ' ;

  if vBoInit then
    begin
    // G_TOTDEBPTP / G_TOTCREPTP
    lStReq := lStReq   + 'CU_DEBIT7 = ' + StrFPoint( vTotDebPTP )
                     + ', CU_CREDIT7 = ' + StrFPoint( vTotCrePTP )
    // G_TOTDEBPTD / G_TOTCREPTD
                     + ', CU_DEBIT8 = ' + StrFPoint( vTotDebPTD )
                     + ', CU_CREDIT8 = ' + StrFPoint( vTotCrePTD ) ;
    end
  else
    begin
    // G_TOTDEBPTP / G_TOTCREPTP
    lStReq := lStReq   + 'CU_DEBIT7 = CU_DEBIT7 + ' + StrFPoint( vTotDebPTP )
                     + ', CU_CREDIT7 = CU_CREDIT7 + ' + StrFPoint( vTotCrePTP )
    // G_TOTDEBPTD / G_TOTCREPTD
                     + ', CU_DEBIT8 = CU_DEBIT8 + ' + StrFPoint( vTotDebPTD )
                     + ', CU_CREDIT8 = CU_CREDIT8 + ' + StrFPoint( vTotCrePTD ) ;
    end ;

  lStReq := lStReq + lStWhere ;

  // ================
  // === MAJ BASE ===
  // ================
  result := ExecuteSQL( lStReq ) = 1 ;

  if not result then
    V_PGI.IoError:=oeSaisie;

end ;


function  CReinitCumulsPointeMS( vCompte : String = '' ; vDossier : String = '' ) : Boolean ;
var lStSQL    : String ;
begin

  result := True ;

  // Uniquement pour les tables partagee en multisoc !
  if not EstTablePartagee( 'GENERAUX' ) then Exit ;

  try
    {JP 28/09/06 : Gestion du multi sociétés}
    lStSQL := 'UPDATE ' + GetTableDossier( vDossier, 'CUMULS' )
                        + ' SET CU_DEBIT6=0, CU_CREDIT6=0, '
                        + 'CU_DEBIT7=0, CU_CREDIT7=0, '
                        + 'CU_DEBIT8=0, CU_CREDIT8=0 ';

    // type de cumuls
    lStSQL := lStSQL + ' WHERE CU_TYPE="' + fbToCumulType( fbGene ) + '"'
                     + ' AND CU_COMPTE2=""' ;
    // pour un compte précis?
    if vCompte <> '' then
        lStSQL := lStSQL + ' AND CU_COMPTE1="' + vCompte +'"' ;

    ExecuteSQL( lStSQL ) ;

  Except
    result := False ;
  end ;

end ;

procedure CRAZCumulsMS( vFB : TFichierBase ; vAxe : string = '' ) ;
var lStWhere : string ;
    lStSQL   : string ;
begin

  // Traitement spécial si appel pour un axe
  if vFb in [fbAxe1..fbAxe5] then
    begin
    vAxe := fbToAxe( vFb ) ;
    vFb  := fbSect ;
    end
  else if vFb <> fbSect then vAxe := '' ;

  // Uniquement pour les tables partagee en multisoc !
  if not EstTablePartagee( fbToTable( vFB ) ) then Exit ;

  // ==== Conditions sur le cumul ====
  // type de cumuls
  lStWhere := lStWhere + ' WHERE CU_TYPE="' + fbToCumulType( vFB ) + '"' ;
  // Axe pour les sections
  if ( vFB = fbSect ) and ( vAxe <> '' )
    then lStWhere := lStWhere + ' AND CU_COMPTE2="' + vAxe +'"' ;

  // Champs MAJ
  lStSQL := 'UPDATE CUMULS SET CU_DEBIT1=0, CU_CREDIT1=0, '
                            + 'CU_DEBIT2=0, CU_CREDIT2=0, '
                            + 'CU_DEBIT3=0, CU_CREDIT3=0, '
                            + 'CU_DEBIT4=0, CU_CREDIT4=0, '
                            + 'CU_DEBIT5=0, CU_CREDIT5=0, '
                            + 'CU_DEBIT9=0, CU_CREDIT9=0, '
                            + 'CU_DEBITAN=0, CU_CREDITAN=0 '
                            + lStWhere ;

  // type de cumuls
  lStSQL := lStSQL + ' WHERE CU_TYPE="' + fbToCumulType( fbGene ) + '"'
                   + ' AND CU_COMPTE2=""' ;

  ExecuteSQL( lStSQL ) ;

end ;


function  CCreateCumulsMS( vFB : TFichierBase ; vCompte, vAxe : String ; vDossier : String = '' ) : Boolean ;
var lStSQLInsert    : string ;
    lStSQLTest      : string ;
begin

  lStSQLTest := 'SELECT CU_COMPTE1 FROM ' + GetTableDossier( vDossier, 'CUMULS')
             +  ' WHERE CU_TYPE="' + fbToCumulType( vFB ) + '"'
                + ' AND CU_COMPTE1="' + vCompte + '"'
                + ' AND CU_COMPTE2="' + vAxe + '"' ;

  // ===================================
  // === Création du cumul si besoin ===
  // ===================================
  result := ExisteSQL( lStSQLTest ) ;
  if not result then
    begin
      try
      lStSQLInsert := 'INSERT INTO ' + GetTableDossier( vDossier, 'CUMULS' )
                  + ' (CU_TYPE, CU_COMPTE1, CU_COMPTE2, CU_EXERCICE, CU_SUITE, CU_ETABLISSEMENT, CU_DEVQTE, CU_QUALIFPIECE,'
                  + ' CU_DEBIT1, CU_DEBIT2, CU_DEBIT3, CU_DEBIT4, CU_DEBIT5, CU_DEBIT6, CU_DEBIT7, CU_DEBIT8, CU_DEBIT9, CU_DEBIT10, CU_DEBIT11, CU_DEBIT12,'
                  + ' CU_CREDIT1, CU_CREDIT2, CU_CREDIT3, CU_CREDIT4, CU_CREDIT5, CU_CREDIT6, CU_CREDIT7, CU_CREDIT8, CU_CREDIT9, CU_CREDIT10, CU_CREDIT11, CU_CREDIT12,'
                  + ' CU_CREDITAN, CU_DEBITAN, CU_QUALIFQTEDIFF )'
                  + ' VALUES ("' + fbToCumulType( vFB ) + '", "' + vCompte + '", "' + vAxe + '", "", "",'
                             + '"", "' + V_PGI.DevisePivot + '", "N",'
                             + ' 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,'
                             + ' 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,'
                             + ' 0, 0, "")' ;
      result := ExecuteSQL( lStSQLInsert ) = 1 ;

      Except
        on E: Exception do
          // Re-test au cas ou l'erreur est due à une duplicate key car crée par un autre user entre-temps...
          // (sera valable après mise en place de l'index unique dans la table CUMULS )
          result := ExisteSQL( lStSQLTest ) ; // Après test, ne fonctionne pas si simplement...En attente
      End ;
    end ;

end ;

end.




