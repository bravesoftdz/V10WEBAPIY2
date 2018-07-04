unit LettUtil;

interface

uses Windows,
     Messages,
     SysUtils,
     Classes,
     Graphics,
     Controls,
     Hqry,
     Ent1,
     HEnt1,
     hmsgbox,
     dialogs,
     HCtrls,
     Forms,
{$IFDEF VER150}
     variants,
{$ENDIF}
     StdCtrls,
     ComCtrls,
     ParamSoc,
{$IFDEF EAGLCLIENT}
{$ELSE}
     DBGrids,
     DB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF COMPTA}
     uLibEcriture,
     ULibWindows,
{$ENDIF}
     UTob,
     SaisUtil,
     SaisComm,
     UtilPGI,
     {$IFNDEF EAGLSERVER}
     LettAuto,
     {$ENDIF}
     {$IFDEF MODENT1}
     CPProcGen,
     CPProcMetier,
     {$ENDIF MODENT1}
     HCompte
     ;

type TCR = Array[1..3] of String17 ;
type REGJAL = RECORD
              D,C        : TCR ;
              Journal    : String3 ;
              Facturier  : String3 ;
              ModeSaisie : String3 ;
              END ;

Type T_LETTECHANGE = RECORD
                     RLL : RLETTR ;
                     DEV : RDEVISE ;
                     Client : boolean ;
                     NbSol  : integer ;
                     Titre  : String ;
                     END ;

Type PG_Lettrage = (pglRegul,pglEcart,pglConvert,pglInverse) ;

Type T_ECARTCHANG = RECORD
                    Cpte     : String17 ;
                    Regul    : PG_Lettrage ;
                    Montant1,Montant2  : Double ;
                    Ref,Lib,RefLettrage : String35 ;
                    Quotite  : double ;
                    Decimale : integer ;
                    DPMin,DPMax : TDateTime ;
                    END ;

Type T_RAPPAUTO = Class
                  General,Auxiliaire : String17 ;
                  Devise             : String3 ;
                  NbC,NbD            : integer ;
                  TotalD,TotalC      : double ;
                  CodeA,CodeZ        : String4 ;
                  END ;

type TL_Rappro = Class
                 General,Auxiliaire,NumTraChq  : String17 ;
                 ChampMarque,ValeurMarque : String ;
                 DebitCur,CreditCur,CouvCur : double ;
                 Debit,Credit,Couv      : double ;
                 DebDev,CredDev,CouvDev : double ;
                 DateC,DateE,DateR      : TDateTime ;
                 RefI,RefE,RefL,Lib     : String35 ;
                 Numero                 : Longint ;
                 Jal,Nature,Exo,CodeD   : String3 ;
                 NumLigne,NumEche       : integer ;
                 CodeL                  : String ;
                 TauxDev                : Double ;
                 Solution,Decim         : integer ;
                 Facture,Client,EditeEtatTva : boolean ;
                 //LG* nouveau champs 06/03/2002
                 SoldePro               : double;  // solde progressif ( sert pour le lettrage par solde )
                 NumLigneDep            : integer; // numéro de la ligne de depart
                 EtatLettrage           : string3; // champs E_ETATLETTRAGE
                 Etablissement          : string3; // champs E_ETABLISSEMENT
                 Etat                   : string; // champs E_ETAT
                 EcartRegulDevise       : double;
                 EcartChangeEuro        : double;
                 EcartChangeFranc       : double;
                 ConvertEuro            : double;
                 ConvertFranc           : double; // ecart de change en euro (!)
                 ModeRappro             : char;  // A mode automatique M mode manuel
                 DatePaquetMax          : TDateTime;
                 DatePaquetMin          : TDateTime;
                 TypeLigne              : integer ; // sert pour le lettrage par anciennete
                 Dossier                : string ;  // SBO 20/02/2006 : lettrage multi-dossier
                 END ;

RINFOSLETT = RECORD
             RefFC,RefRC,RefFF,RefRF,DateFC,DateRC,DateFF,DateRF : String3 ;
             EgalC,EgalF     : boolean ;
             TolC,TolF       : integer ;
             //LG* 10/11/2001
             ChoixValid      : string3; // action a effecter sur le F10 ds le lettrage AL1 : solde automatique AL2 : lettrage partiel Al3 : ecriture simplifie
             EcartDebit      : double;  // ecart au debit pour une ecriture de regul
             EcartCredit     : double;  // ecart au credit pour une ecriture de regul
             GeneralDebit    : string;  // compte general au debit par defaut
             GeneralCredit   : string;  // compte general au debit par defaut
             Journal         : string;  // journal par defaut
             Libelle         : string;  // libelle par defaut
             AvecDebitCredit : boolean; // si true le lettrage auto doit avois des mvt au debit et credit
             LetTotal        : boolean; // si true seul le lettrage total est genere en mode automatique
             AvecLetPartiel  : boolean; // si true les ecritures partiellement lettrees sont prise ne compte dans le lettrage auto
             ModeLettrage    : string;  // Si PCL : lettrage dans deux grilles avec solde progressif
             MaxProf         : integer; // profondeur max du lettrage combinatoire en mode automatique
             MaxDuree        : integer; // duree max pour le lettrage automatique
             JournalChange   : string;  // journal de change par defaut
             CAvecOD         : boolean ;
             END ;

Tisf = (isfCli,isfFou,isFDiv) ;

CONST SumDC      = '(E_DEBIT+E_CREDIT+E_DEBITDEV+E_CREDITDEV)' ;
      SumD       = '(E_DEBIT+E_DEBITDEV)' ;
      SumC       = '(E_CREDIT+E_CREDITDEV)' ;
      DSurDC     = '(E_DEBIT+E_DEBITDEV)/(E_DEBIT+E_CREDIT+E_DEBITDEV+E_CREDITDEV)' ;
      CSurDC     = '(E_CREDIT+E_CREDITDEV)/(E_DEBIT+E_CREDIT+E_DEBITDEV+E_CREDITDEV)' ;
      SumSolde   = '(E_DEBIT-E_CREDIT+E_DEBITDEV-E_CREDITDEV)' ;
      SoldeSurDC = '(E_DEBIT-E_CREDIT+E_DEBITDEV-E_CREDITDEV)/(E_DEBIT+E_CREDIT+E_DEBITDEV+E_CREDITDEV)' ;

{$IFNDEF EAGLSERVER}
procedure InitTablesLibresTiers ( TT : TTabSheet ) ;

function  FormatCheque(sr,nc : string ) : string ;
{$IFDEF COMPTA}
{$IFDEF EAGLCLIENT}
procedure ChangeAspectLettre (GL : THGrid);
{$ELSE}
procedure ChangeAspectLettre ( GL : TDBGrid ) ;
function  GetRefParamDB ( T : TDataSet ; CRef : String ) : String ;
{$ENDIF}
procedure RempliSelectEuro ( C : THValComboBox ; TitreEuro : String ) ;
{$ENDIF}
Function  SelectQDL ( Quoi : Integer ; St : String = '') : String ;
function  DateCorrecte ( DD : TDateTime ) : integer ;
procedure AjouteDevise ( CDevise : THValComboBox ; CodeD : String ) ;
//LG*
procedure CAssignListRappro( G : THGrid ; LD,LC : TList ; vStModeRappro : Char ; vBoLettrageDevise : boolean);
function  COBMVersRappro( OBM : TOBM; vIndex : integer ; vBoLettrageDevise : boolean ) : TL_Rappro;
function  COBMVersRappAuto( OBM : TOBM ) : T_RAPPAUTO;
{$IFDEF COMPTA}
function  CTHGridVersTLett( G : THGrid ; vBoLettrageDevise : boolean ; EnrSelect : boolean = false; EcrLett : boolean = false) : TList;
{$ENDIF}
procedure CTLettVersTHGrid( G : THGrid ; TLett : TList; RL : RLETTR;GrilleDebit : boolean );
function  ChargeInfosLett : RINFOSLETT ;
function  LettrerReference ( RL : RLETTR ; TLett : TList ; CAvecOD,CODSALFOU : boolean;IFL : RINFOSLETT;GeneNumSol : integer ; var CodeL : string) : boolean;
procedure FichierLettrage ( RL : RLETTR ; TLett : TList;NoSol : integer ; Total : boolean ; ModeRappro : Char ; var CodeL : string ; EcartPivot,EcartCV,EcartDev : double );
procedure LettrerSoldeProgressif(RL : RLETTR ; TLett : TList ; Var CodeL : string);
procedure LettrerSolde(RL : RLETTR ; TLett : TList ; var Codel : string ) ;
procedure CLettrerMontant(RL : RLETTR ; TLett : TList ; CAvecOD : boolean ; MaxProf,MaxDuree : integer ; Var CodeL : string) ;
function  MemeRef ( LaRef,RefTest : String ; Client : boolean; IFL : RINFOSLETT ) : Boolean ;
function  CSuppLett(PasEcart,AvecRegul,EcartChange,AvecConvert : boolean) : Char ;
procedure CTOBVersTHGrid ( G : THGrid ; TOBLigne : TOB ; Devise : boolean ; CRef : string = '') ;
procedure AnalyseTitreMontant ( Var Titre : String ; Devise : boolean ) ;
function  CGetTOBSelectFromGrid( G : THGrid ; vBoTotal : boolean=false) : TOBM;
function  BonneDateRef ( LaDate,DateTest : TDateTime ; Client : boolean ; IFL : RINFOSLETT) : Boolean ;
Function  isISF ( L : TL_Rappro ; CODSALFOU : boolean) : Tisf ;
function  QuelleRef ( L : TL_Rappro ; CAvecOD,CODSALFOU : boolean; IFL : RINFOSLETT) : String ;
function  QuelleDate ( L : TL_Rappro ; Var ControleDate : boolean ; CAvecOD,CODSALFOU : boolean;IFL : RINFOSLETT ) : TDateTime ;
Function  BonneNature ( Nat1,Nat2 : String ; CAvecOD : boolean) : boolean ;
procedure CAjouteChampsSuppLett ( O : TOB );
function  EstSelect ( G : THGrid ; Lig : integer ) : boolean ;
Procedure ConvertCouverture ( O : TOBM ; Origine : TSorteMontant ) ;
function  TripoteSt( St : String ) : String ;
function  GoReqMajLet ( O : TOBM ; Gene,Auxi : String17 ; Stamp : TDateTime ; OKL : boolean ) : boolean ;
function  LWhereBase ( OkLettrage,OkTiers,SaisieTreso,SansLesPartiels : boolean ; LettrageEnSaisie : boolean = false ; CritEtatLettrage : string = '' ) : String ;
function  LWhereComptes ( F : TForm ) : String ;
//function  LWhereComptes2 ( F : TForm ) : String ;  FB 19757 - suppresion de cette fonction, on utilise LWhereCompte
function  LWhereVide : String ;
function  LTri : String ;
procedure RemplirInfosRegul ( Var J : REGJAL ; Value : String );

//function  GetSetCodeLettre ( RL : TL_Rappro ; DernLettrage : string = '' ) : String ;
function  GetSetCodeLettre ( vStGen , vStAux : string ; vStDernLettrage : string = '') : string ;
Function  Fourchette ( Champ,F1,F2 : String ; tft : TFieldType ) : String ;
function  BonSigne ( Reste,ValTest : double ; Signe : TSigne ) : boolean ;
function  PasTouche ( St : String ) : boolean ;
procedure FormatLettrage ( C : THNumEdit ; CDevise : THValComboBox ; CodeD : String ) ;
Function  RazTvaEnc ( X : TL_Rappro ) : boolean ;
procedure CEtudieModeLettrage ( var R : RLETTR ) ;
procedure CExecDelettreMvtExoRef( vStGeneral, vStAux : string );
procedure CExecDeLettrage( vStGeneral, vStAux : string );
{$IFDEF COMPTA}
procedure CDeLettrageMulti ( G : THGrid ; vStPrefixe : string = 'G_' ; vBoMvtSurExo : boolean = false ) ;
{$ENDIF}
function  CLettrageParAnc( vBoVentilDebit : boolean ; RL : RLETTR ; TLett : TList ; RDEV : RDEVISE ; vB : T_RAPPAUTO = nil ) : boolean ;


function  NextLigneLettre ( L : TList ; FromLig : integer ; Var Reste : double ; ValTest : double ; Signe : TSigne ) : integer ;
procedure RapprocheLesLignes ( L1,L2 : TList ; i1,i2 : integer ) ;
function  TrouveLesQuels ( LD,LC : TList ; Var L1,L2 : TList ; Var i1,i2 : integer ) : boolean ;
procedure RefEpuise ( LD,LC : TList ) ;

//function  LettrageTotalTOB ( vTOBEcr : TOB ) : boolean ;
{$ENDIF EAGLSERVER}

function  CodeSuivant ( CL : String ) : String ;
{$IFNDEF EAGLSERVER}
function  CControlePresenceLettrage ( vTOB : TOB ) : boolean ;
{$IFDEF COMPTA}
procedure CLettrageTOB ( vTOBEcr : TOB ) ;
{$ENDIF}
{$ENDIF}
procedure DelettreEcriture (var MM : RMVT ; TOBECR : TOB);
procedure DelettreEcritureinMem (TOBECR : TOB);
implementation

{$IFNDEF EAGLSERVER}

Function RazTvaEnc ( X : TL_Rappro ) : boolean ;
BEGIN
Result:=False ;
if Not VH^.OuiTvaEnc then Exit ;
if ((X.Nature='FC') or (X.Nature='AC') or (X.Nature='FF') or (X.Nature='AF')) then Exit ;
if EstJalFact(X.Jal) then Exit ;
if X.EditeEtatTva then Exit ;
if Not EstCollFact(X.General) then Exit ;
Result:=True ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : LG* Laurent GENDREAU
Créé le ...... : 13/12/2001
Modifié le ... : 11/09/2002
Description .. : -Remplacement des Fields[0] pas des FindField pour eAGL
Suite ........ : -Suppression des Q.Edit et Q.Post pour l'eAGL
Suite ........ :
Suite ........ : LG - 08/04/2002 - fonction deplacer de lettrage.pas
Suite ........ :
Suite ........ : - LG - 11/09/2002 - fermeture de la requete au plus tot
Mots clefs ... :
*****************************************************************}
{function GetSetCodeLettre ( RL : TL_Rappro ; DernLettrage : string = '') : String ;
Var Q : TQuery ;
    OldLet,NewLet : String ;
    lBoQEOF : boolean ;
BEGIN

NewLet:='' ;
if RL.Auxiliaire<>'' then Q:=OpenSQL('Select T_DERNLETTRAGE from TIERS Where T_AUXILIAIRE="'+RL.Auxiliaire+'"',true)
                     else Q:=OpenSQL('Select G_DERNLETTRAGE from GENERAUX Where G_GENERAL="'+RL.General+'"',true) ;
if RL.Auxiliaire<>'' then NewLet:=Q.FindField('T_DERNLETTRAGE').AsString else NewLet:=Q.FindField('G_DERNLETTRAGE').AsString ;
lBoQEOF:=Q.EOF ; Ferme(Q) ;
try
if Not lBoQEOF then
   BEGIN
   if DernLettrage='' then OldLet:=NewLet else OldLet:=DernLettrage ;
   if NewLet=OldLet then
      BEGIN
      DernLettrage:=CodeSuivant(NewLet) ;
      if RL.Auxiliaire<>'' then ExecuteSQL('update TIERS set T_DERNLETTRAGE="'+DernLettrage+'" Where T_AUXILIAIRE="'+RL.Auxiliaire+'"')
      else ExecuteSQL('update GENERAUX set G_DERNLETTRAGE="'+DernLettrage+'" Where G_GENERAL="'+RL.General+'"')
      END else V_PGI.IoError:=oeUnknown ;
   END else V_PGI.IoError:=oeUnknown ;
result:=DernLettrage;
except
 on E:Exception do
  begin
   MessageAlerte('Problème lors de la recherche du code lettre'+#10#13+E.message) ;
   raise ;
  end;
end;
END ;
}

function PasTouche ( St : String ) : boolean ;
BEGIN
 if Length(St)>=4 then PasTouche:=((St[1]='X') or (St[2] in ['X','#'])) else PasTouche:=False ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 09/10/2002
Modifié le ... :   /  /
Description .. : - 09/10/2002 - on ne retourne que les 4 premiers caractères
Suite ........ : pour ne pas plante si le dernier code lettre est ZZZZ. ( cas
Suite ........ : des recup )
Mots clefs ... :
*****************************************************************}

function BonSigne ( Reste,ValTest : double ; Signe : TSigne ) : boolean ;
BEGIN
BonSigne:=False ;
Case Signe of
   sgEgal   : if Abs(Reste+ValTest)<>Abs(Reste)+Abs(ValTest) then Exit ;
   sgOppose : if Abs(Reste+ValTest)=Abs(Reste)+Abs(ValTest) then Exit ;
   END ;
BonSigne:=True ;
END ;

{$IFDEF COMPTA}
{$IFDEF EAGLCLIENT}
procedure ChangeAspectLettre (GL : THGrid);
Var
  i : integer ;
BEGIN
  for i:=0 to GL.ColCount-1 do BEGIN
    if i in [3,4] then GL.ColColors[i] := clNavy;
    if i in [5,6] then GL.ColColors[i] := clMaroon;
    if i in [7]   then GL.ColColors[i] := clGreen;
  END;
END;
{$ELSE}
procedure ChangeAspectLettre ( GL : TDBGrid ) ;
Var i : integer ;
BEGIN
CentreDBGrid(GL) ;
for i:=0 to GL.Columns.Count-1 do
    BEGIN
    if i in [3,4] then GL.Columns.Items[i].Font.Color:=clNavy ;
    if i in [5,6] then GL.Columns.Items[i].Font.Color:=clMaroon ;
    if i in [7] then GL.Columns.Items[i].Font.Color:=clGreen ;
    END ;
END ;
{$ENDIF}

{$ENDIF}

Function  Fourchette ( Champ,F1,F2 : String ; tft : TFieldType ) : String ;
Var St1,St2 : String ;
BEGIN
case tft of
  ftFloat :
           BEGIN
           if F1<>'' then F1:=StrfPoint(Valeur(F1)) ;
           if F2<>'' then F2:=StrfPoint(Valeur(F2)) ;
           END ;
  END ;
St1:='' ; St2:='' ;
if F1<>'' then BEGIN if tft=ftString then F1:='"'+F1+'"' ; St1:=Champ+'>='+F1 ; END ;
if F2<>'' then BEGIN if tft=ftString then F2:='"'+F2+'"' ; St2:=Champ+'<='+F2 ; END ;
if ((St1<>'') and (St2<>'')) then St2:=' AND '+St2 ;
Fourchette:=St1+St2 ;
END ;

procedure FormatLettrage ( C : THNumEdit ; CDevise : THValComboBox ; CodeD : String ) ;
Var St,Symbole : String ;
    Decim      : integer ;
BEGIN
if CDevise.Items.Count<=1 then Exit ; CDevise.Value:=CodeD ;
St:=CDevise.Text ; if St='' then Exit ;
Symbole:=Trim(ReadtokenSt(St)) ; Decim:=StrToInt(Trim(ReadTokenSt(St))) ;
if ((C.CurrencySymbol<>Symbole) or (C.Decimals<>Decim)) then ChangeMask(C,Decim,Symbole) ;
END ;

procedure AjouteDevise ( CDevise : THValComboBox ; CodeD : String ) ;
Var R : RDEVISE ;
BEGIN
if CDevise.Values.IndexOf(CodeD)>=0 then Exit ;
if CodeD=V_PGI.DevisePivot then
   BEGIN
   CDevise.Values.Add(V_PGI.DevisePivot) ; CDevise.Items.Add(V_PGI.SymbolePivot+';'+IntToStr(V_PGI.OkDecV)+';') ;
   END else
   BEGIN
   R.Code:=CodeD ; GetInfosDevise(R) ;
   CDevise.Values.Add(CodeD) ; CDevise.Items.Add(R.Symbole+';'+IntToStr(R.Decimale)+';') ;
   END ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... : 25/05/2004
Description .. : GCO - 25/05/2004 Sortie de la fonction du contexte IFDEF COMPTA
Suite ........ : pour les standards Compta PGI
Mots clefs ... : 
*****************************************************************}
procedure InitTablesLibresTiers ( TT : TTabSheet ) ;
Var LesLib : HTStringList ;
    i : Integer ;
    St,Titre,Ena : String ;
    Trouver : Boolean ;
    LL      : TLabel ;
    CC      : THCpteEdit ;
BEGIN
if TT=Nil then Exit ; Trouver:=False ;
LesLib:=HTStringList.Create ; GetLibelleTableLibre('T',LesLib) ;
for i:=0 to LesLib.Count-1 do
    BEGIN
    St:=LesLib.Strings[i] ; Titre:=ReadTokenSt(St) ; Ena:=St ;
    LL:=TLabel(TForm(TT.Owner).FindComponent('TT_TABLE'+IntToStr(i))) ;
    CC:=THCpteEdit(TForm(TT.Owner).FindComponent('T_TABLE'+IntToStr(i))) ;
    if LL<>Nil then
       BEGIN
       LL.Caption:=Titre ; LL.Enabled:=(Ena='X') ;
       if ((EstSerie(S3)) and (i>2)) then LL.Visible:=False ;
       if CC<>Nil then
          BEGIN
          CC.Enabled:=LL.Enabled ; if CC.Enabled then Trouver:=True ;
          if ((EstSerie(S3)) and (i>2)) then CC.Visible:=False ;
          END ;
       END ;
    END ;
TT.TabVisible:=Trouver ;
LesLib.Clear ; LesLib.Free ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... : 04/06/2007
Description .. : 
Suite ........ : GCO - 25/05/2004 Sortie de la fonction du contexte IFDEF 
Suite ........ : COMPTA
Suite ........ : pour les standards Compta PGI
Suite ........ : - LG - FB 19757 - suppression de la condition sur les auxi 
Suite ........ : vide
Mots clefs ... : 
*****************************************************************}
Function LWhereComptes ( F : TForm ) : String ;
Var StAuxi,StGene,St,StD : String ;
    G1,G2,A1,A2 : TEdit ;
    D           : THValComboBox ;
BEGIN
St:='' ; StD:='' ;
G1:=TEdit(F.FindComponent('E_GENERAL')) ; G2:=TEdit(F.FindComponent('E_GENERAL_')) ;
A1:=TEdit(F.FindComponent('E_AUXILIAIRE')) ; A2:=TEdit(F.FindComponent('E_AUXILIAIRE_')) ;
StGene:=Fourchette('E_GENERAL',G1.Text,G2.Text,ftString) ;
StAuxi:=Fourchette('E_AUXILIAIRE',A1.Text,A2.Text,ftString) ;
if StGene<>'' then
   BEGIN
   if StAuxi<>'' then St:=StAuxi+' AND '+StGene else St:=StGene ; //+' AND E_AUXILIAIRE=""' ; FB 19757
   END else if StAuxi<>'' then St:=StAuxi ;
if St<>'' then St:=' AND ('+St+')' ;
D:=THValComboBox(F.FindComponent('E_DEVISE')) ;
if D<>Nil then BEGIN if D.Value<>'' then StD:=' AND E_DEVISE="'+D.Value+'"' ; END ;
LWhereComptes:=St+StD ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... :   /  /
Modifié le ... : 15/09/2004
Description .. :
Mots clefs ... :
*****************************************************************}
{Function LWhereComptes2 ( F : TForm ) : String ;
Var StAuxi,StGene,St,StD : String ;
    G1,G2,A1,A2 : TEdit ;
    D           : THValComboBox ;
BEGIN
St:='' ; StD:='' ;
G1:=TEdit(F.FindComponent('E_GENERAL')) ; G2:=TEdit(F.FindComponent('E_GENERAL_')) ;
A1:=TEdit(F.FindComponent('E_AUXILIAIRE')) ; A2:=TEdit(F.FindComponent('E_AUXILIAIRE_')) ;
StGene:=Fourchette('E_GENERAL',G1.Text,G2.Text,ftString) ;
StAuxi:=Fourchette('E_AUXILIAIRE',A1.Text,A2.Text,ftString) ;
if StGene<>'' then
   BEGIN
   if StAuxi<>'' then St:=StAuxi+' AND '+StGene else St := StGene;
   END else if StAuxi<>'' then St:=StAuxi ;
if St<>'' then St:=' AND ('+St+')' ;
D:=THValComboBox(F.FindComponent('E_DEVISE')) ;
if D<>Nil then BEGIN if D.Value<>'' then StD:=' AND E_DEVISE="'+D.Value+'"' ; END ;
LWhereComptes2:=St+StD ;
END ; }

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... : 25/05/2004
Suite ........ : GCO - 25/05/2004 Sortie de la fonction du contexte IFDEF COMPTA
Suite ........ : pour les standards Compta PGI
Mots clefs ... :
*****************************************************************}
function LWhereVide : String ;
BEGIN
LWhereVide:='E_JOURNAL="ZZZ" AND E_EXERCICE="ZZZ" AND E_DATECOMPTABLE="'+USDateTime(IDate1900)+'" AND E_NUMEROPIECE=-1' ;
END ;

function FormatCheque(sr,nc : string ) : string ;
begin
result:=sr+' '+nc ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 04/05/2004
Modifié le ... : 21/07/2004
Description .. : - LG - 04/05/2004 - Modif pour l'eAgl ( module LetRegul)
Suite ........ : - BPY - 21/07/2004 - modif pour faire marché !
Suite ........ : ajout d'un "solde" dans le cas des ecart de change sinon
Suite ........ : ca plante dans letregul !
Mots clefs ... :
*****************************************************************}
Function SelectQDL ( Quoi : Integer ; St : String = '') : String ;
BEGIN
Case Quoi of
{Ecart}  0 :
             Result:='Select E_GENERAL, E_AUXILIAIRE, Max(T_LIBELLE), '
             +'Sum('+DSurDC+') SDSURDC, Sum(E_COUVERTURE*'+DSurDC+') CDSURDC, '
             +'Sum('+CSurDC+') SCSurDC, Sum(E_COUVERTURE*'+CSurDC+') CCSurDC, '
             +'Sum(E_COUVERTURE*'+SoldeSurDC+') SOLDE, '
             +'E_LETTRAGE, E_DEVISE, Max(E_LETTRAGEDEV)   '
             +'From Ecriture Left Outer Join Tiers On E_AUXILIAIRE=T_AUXILIAIRE ' ;

{Delett} 2 : Result:='Select E_GENERAL, E_AUXILIAIRE, Max(T_LIBELLE) , '
             +'Sum('+DSurDC+') DSurDC, Sum(E_DEBIT) , '
             +'Sum('+CSurDC+') CSurDC, Sum(E_CREDIT) , Sum(E_DEBIT-E_CREDIT) SOLDE, '
             +'E_LETTRAGE, Max(E_DEVISE) , Max(E_LETTRAGEDEV)  '
             +'From Ecriture Left Outer Join Tiers On E_AUXILIAIRE=T_AUXILIAIRE ' ;
{Regul}  1 :
             Result:='Select E_GENERAL, E_AUXILIAIRE, Max(T_LIBELLE) , '
             +'Sum('+DSurDC+')  DSurDC, Sum(E_DEBIT) , '
             +'Sum('+CSurDC+') CSurDC, Sum(E_CREDIT) , Sum(E_DEBIT-E_CREDIT) SOLDE, '
             +'E_LETTRAGE, E_DEVISE, Max(E_LETTRAGEDEV)  '
             +'From Ecriture Left Outer Join Tiers On E_AUXILIAIRE=T_AUXILIAIRE ' ;

{Convert}3 : Result:='Select E_GENERAL, E_AUXILIAIRE, Max(T_LIBELLE) , '
             +'Sum('+DSurDC+') DSurDC, Sum(E_COUVERTURE*'+DSurDC+') CDSurDC, 0, ' //Sum(E_COUVERTUREEURO*'+DSurDC+'), '
             +'Sum('+CSurDC+') CSurDC, Sum(E_COUVERTURE*'+CSurDC+') CCSurDC, 0, ' //Sum(E_COUVERTUREEURO*'+CSurDC+'), '
             +'Sum(E_COUVERTURE*'+SoldeSurDC+') SoldeSurDC, 0, ' //Sum(E_COUVERTUREEURO*'+SoldeSurDC+'), '
             +'E_LETTRAGE, E_DEVISE, Max(E_LETTRAGEDEV) , "-", ' //Max(E_LETTRAGEEURO) '
             +'From Ecriture Left Outer Join Tiers On E_AUXILIAIRE=T_AUXILIAIRE ' ;
   END ;
If St<>'' Then Result:=Result+' '+St+' ' ;
END ;

function DateCorrecte ( DD : TDateTime ) : integer ;
Var Err : integer ;
BEGIN
Err:=0 ;
if DD<VH^.Encours.Deb then Err:=1 else
 if ((VH^.Suivant.Code<>'') and (DD>VH^.Suivant.Fin)) then Err:=2 else
  if ((VH^.Suivant.Code='') and (DD>VH^.Encours.Fin)) then Err:=2 else
   if DD<=VH^.DateCloturePer then Err:=3 else
{   if DD<=V_PGI.DateClotureDef then Err:=4} ;
Result:=Err ;
END ;

{$IFDEF COMPTA}
procedure RempliSelectEuro ( C : THValComboBox ; TitreEuro : String ) ;
BEGIN
C.Items.Clear ; C.Values.Clear ;
C.Values.Add(V_PGI.DevisePivot) ; C.Items.Add(RechDom('TTDEVISETOUTES',V_PGI.DevisePivot,False)) ;
//C.Values.Add('ECU') ; C.Items.Add(RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False)) ;
C.Values.Add('MIX') ; C.Items.Add(RechDom('TTDEVISETOUTES',V_PGI.DevisePivot,False)+' + '+RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False)) ;
C.Value:=V_PGI.DevisePivot ;
END ;
{$ENDIF}

{$IFNDEF EAGLCLIENT}
function  GetRefParamDB ( T : TDataSet ; CRef : String ) : String ;
Var Nam : String ;
    FF : TField ;
BEGIN
Result:='ERROR' ;
if Not T.Active then Exit ;
if CRef='' then Exit ; Nam:='' ;
if CRef='REF' then Nam:='E_REFINTERNE' else
if CRef='LIB' then Nam:='E_LIBELLE' else
if CRef='REX' then Nam:='E_REFEXTERNE' else
if CRef='RLI' then Nam:='E_REFLIBRE' else
if CRef='NUM' then Nam:='E_NUMEROPIECE' else
if ((CRef>='TL0') and (CRef<='TL9')) then Nam:='E_TEXTELIBRE'+CRef[3] else
if ((CRef>='CL0') and (CRef<='CL3')) then Nam:='E_TABLE'+CRef[3] else ;
if Nam='' then Exit ;
FF:=T.FindField(Nam) ;
if FF=Nil then Exit ;
Result:=FF.AsString ;
END ;
{$ENDIF}

//LG* 16/11/20001

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 16/11/2001
Modifié le ... : 03/11/2004
Description .. : 1- Parcours la grille passé en parametre,
Suite ........ : 2- Pour les lignes selectionnées recupere les obm associés,
Suite ........ : les transformes en TL_Rappro  et remplie les TList LD ou LC
Suite ........ : en fonction du sens du mouvement.
Suite ........ : LD : TList des debits
Suite ........ : LC : TList des credits
Suite ........ :
Suite ........ : - 12/02/2003 - FB 10490 - les debits peuvent etre negatif,
Suite ........ : utilisation de la fct Abs pour determiner la bonne liste
Suite ........ : - LG - 03/11/2004 - FB 14765 - correction pour le lettrage
Suite ........ : en devise
Mots clefs ... :
*****************************************************************}
procedure CAssignListRappro( G : THGrid ; LD,LC : TList ; vStModeRappro : Char ; vBoLettrageDevise : boolean );
var
 j     : integer;
 L     : TL_Rappro ;
 O     : TOBM;
begin
 if (G=nil) or (LD=nil) or (LC=nil) then exit;
 for j:=1 to G.RowCount-1 do
  begin
  if EstSelect(G,j) then
   begin
    O:=GetO(G,j); // recupere l'OBM associé à la ligne
    L:=COBMVersRappro(O,j,vBoLettrageDevise); // transforme OBM en TL_Rappro
    L.ModeRappro:=vStModeRappro ;
    if Abs(L.DebitCur)>0 then LD.Add(L) else LC.Add(L); // ajout à la bonne TList en fonction du sens
   end; // if
  end; // for
end;

{***********A.G.L.***********************************************
Auteur  ...... : Lauretn GENDREAU
Créé le ...... : 13/02/2003
Modifié le ... : 03/11/2004
Description .. : - 13/02/2003 - FB 10219 - on controle que l'on peut
Suite ........ : selectionner la ligne
Suite ........ : - LG - 03/11/2004 - FB 14765 - correction pour le lettrage
Suite ........ : en devise
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
function CTHGridVersTLett( G : THGrid ; vBoLettrageDevise : boolean ; EnrSelect : boolean = false; EcrLett : boolean = false) : TList;
var
 j     : integer;
 L     : TL_Rappro ;
 O     : TOBM;
 TLett : TList;
begin
 result:=nil ; if (G=nil) then exit ; TLett := TList.Create ;
 for j:=1 to G.RowCount-1 do
  begin
  if ( not EnrSelect and EstSelect(G,j) ) or EnrSelect then
   begin
    if not CIsRowLock(G,j) then
     begin
      O:=GetO(G,j); // recupere l'OBM associé à la ligne
      L:=COBMVersRappro(O,j,vBoLettrageDevise); // transforme OBM en TL_Rappro
      if L<>nil then TLett.Add(L);
     end; // if
   end; // if
  end; // for
 result := TLett;
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 08/03/2002
Modifié le ... : 08/03/2002
Description .. :  code = existence d'une pièce de régularisation qui n'existe
Suite ........ : pas encore et qui sera générée au moment de la validation
Suite ........ : du lettrage.
Suite ........ :
Suite ........ : code  + existence d'une pièce d'écart de change qui
Suite ........ : n'existe pas encore et qui sera générée au moment de la
Suite ........ : validation du lettrage.
Suite ........ :
Suite ........ : code @ existence d'une pièce d'écart de conversion qui
Suite ........ : n'existe pas encore et qui sera générée au moment de la
Suite ........ : validation du lettrage.
Suite ........ :
Suite ........ : code  ? lettrage nécessitant une pièce d'écart de change
Suite ........ : non générée.
Suite ........ :
Suite ........ : code  # existence d'une pièce de régularisation et d'une
Suite ........ : pièce d'écart de change qui n'existent pas encore et qui
Suite ........ : seront générées au moment dela validation du lettrage.
Suite ........ :
Suite ........ : code  × existence d'une pièce de régularisation et d'une
Suite ........ : pièce d'écart de conversion qui n'existent pas encore et qui
Suite ........ : seront générées au moment de la validation du lettrage.
Suite ........ :
Suite ........ : code  * existence d'une pièce de régularisation et
Suite ........ : nécessitant une pièce d'écart de change non générée.
Suite ........ :
Mots clefs ... :
*****************************************************************}
function CSuppLett(PasEcart,AvecRegul,EcartChange,AvecConvert : boolean) : Char ;
BEGIN
if AvecRegul then
   BEGIN
   if EcartChange then Result:='#' else
      BEGIN
      if PasEcart then
         BEGIN
         if AvecConvert then Result:='§'{ne devrait pas exister} else Result:='*' ;
         END else
         BEGIN
         if AvecConvert then Result:='¤' else Result:='=' ;
         END ;
      END ;
   END else
   BEGIN
   if EcartChange then Result:='+' else
      BEGIN
      if PasEcart then
         BEGIN
         if AvecConvert then Result:='!'{ne devrait pas exister} else Result:='?' ;
         END else
         BEGIN
         if AvecConvert then Result:='@' else Result:=' ' ;
         END ;
      END ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 17/06/2002
Modifié le ... : 22/01/2004
Description .. : - 17/06/2002 - affectation des champs e_lettragedev et
Suite ........ : e_lettrageeuro
Suite ........ : -03/06/2003 - FB 12460 - on ne mets plus a jour les
Suite ........ : ecritures qui n'ont pas ete lettres
Suite ........ : - 12/01/2004 - FB 12460 - correction d'un erreur de portage 
Suite ........ : entre VDEV et V421 ( voir com ds le code )
Suite ........ : -22/01/04 - en lettrage auto, on sort du traitement pour les
Suite ........ : ecritures solution:=0
Mots clefs ... : 
*****************************************************************}
procedure CTLettVersOBM( L : TL_Rappro ; O : TOBM ; RL : RLETTR ; GrilleDebit : boolean ; AvecValSupp : boolean = true );
var
 C : double;
 AvecConvert,EcartChange : boolean;
begin
 if ((L.DebitCur=0) and GrilleDebit) or ((L.CreditCur=0) and not GrilleDebit) then exit;
 // en lettrage auto
 if (Trim(L.CodeL)='') and (trim(L.ModeRappro)<>'M') then exit;
 if (L.Solution = 0) and (trim(L.ModeRappro)<>'M') then exit ; // 12/01/2004 - suppression du commentaire
 // affectation du code lettre supplementaire
 AvecConvert:=(L.ConvertEuro<>0) or (L.ConvertFranc<>0);
 EcartChange:=(L.EcartChangeEuro<>0) or (L.EcartChangeFranc<>0);
 if (L.EtatLettrage='TL') and (length(L.CodeL)=4) then // on affecte le code supplementaire en lettrage uniquement pour les lettrages totals
  L.CodeL:=L.CodeL+CSuppLett(false,false,EcartChange,AvecConvert) ;
 if trim(L.ModeRappro) <> 'M' then
  begin
   // on mode manuel, le code lettrage est affecte par la fonction FinirLettrage
   O.PutValue('E_LETTRAGE',L.CodeL) ; O.PutValue('E_ETATLETTRAGE',L.EtatLettrage) ;
   O.TraiteLeEtat(false,false,EcartChange,AvecConvert,True,0,L.ModeRappro) ;
   O.PutValue('E_DATEPAQUETMAX',L.DatePaquetMax) ; O.PutValue('E_DATEPAQUETMIN',L.DatePaquetMin) ;
  end; // if
 if (L.EtatLettrage='TL') then
  begin
   O.PutValue('E_COUVERTURE',Arrondi(Abs(O.GetValue('E_DEBIT')-O.GetValue('E_CREDIT')),V_PGI.OkDecV)) ;
   O.PutValue('E_COUVERTUREDEV',Arrondi(Abs(O.GetValue('E_DEBITDEV')-O.GetValue('E_CREDITDEV')),V_PGI.OkDecV)) ;
  end
   else
    begin
     C:=L.CouvCur ;
     if RL.LettrageDevise then
       BEGIN
       O.PutValue('E_COUVERTUREDEV',C) ; O.PutValue('E_LETTRAGEDEV','X') ;
       ConvertCouverture(O,tsmDevise) ;
       END else
       BEGIN
       O.PutValue('E_COUVERTURE',C) ;
       ConvertCouverture(O,tsmPivot) ; O.PutValue('E_LETTRAGEDEV','-') ;
       END ;
   end;
   {Ruses pour gestion interne : les ecarts de regul, conversion sont mis dans des champs inutilise de la TOBM}
   if AvecValSupp then
    begin
     O.PutValue('OLDDEBIT',L.EcartRegulDevise) ;
     O.PutValue('OLDCREDIT',L.EcartChangeEuro) ;
     O.PutValue('RATIO',L.EcartChangeFranc) ;
     O.PutValue('CONVERTEURO',L.ConvertEuro) ;
     O.PutValue('CONVERTFRANC',L.ConvertFranc) ;
    end ;

end;

procedure CTLettVersTHGrid( G : THGrid ; TLett : TList ; RL : RLETTR ; GrilleDebit : boolean );
var
 i : integer;
 O : TOBM;
 L : TL_Rappro;
begin
 for i:=0 to TLett.Count-1 do
  begin
   L:=TL_Rappro(TLett[i]) ;
   O:=GetO(G,L.NumLigneDep) ;
   G.Row:=L.NumLigneDep ;
   CTLettVersOBM(L,O,RL,GrilleDebit) ;
  end; // for
end;



function GetChampRef(CRef : string) : string;
begin
 result:='E_REFINTERNE';
 if CRef='REF' then result:='E_REFINTERNE' else
 if CRef='LIB' then result:='E_LIBELLE' else
 if CRef='REX' then result:='E_REFEXTERNE' else
 if CRef='RLI' then result:='E_REFLIBRE' else
 if CRef='NUM' then result:='E_NUMEROPIECE';
end;

// LG*
procedure AnalyseTitreMontant ( Var Titre : String ; Devise : boolean ) ;
Var TD,TC : boolean ;
    Suff : String ;
BEGIN
{OKOK}
Titre:=uppercase(Trim(Titre)) ;
TD:=Pos('E_DEBIT',Titre)>0 ; TC:=Pos('E_CREDIT',Titre)>0 ; if ((Not TD) and (Not TC)) then Exit ;
if Devise then Suff:='DEV' ;
if TD then Titre:='E_DEBIT'+Suff else Titre:='E_CREDIT'+Suff ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/03/2002
Modifié le ... : 06/03/2002
Description .. : rempli la ligne courante de la grille G avec la tob TOBLigne.
Suite ........ : attention : La grille doit utilise une liste parametre
Suite ........ :                  une colonne supplementaire servant à la
Suite ........ :                   selection
Suite ........ : parametre :
Suite ........ : devise : true si saisie en devise
Suite ........ : oppose : true si saisie en contrevaleur
Suite ........ : CRef : colonne à utilise poiur le champs refence interne
Mots clefs ... :
*****************************************************************}
procedure CTOBVersTHGrid ( G : THGrid ; TOBLigne : TOB ; Devise : boolean ; CRef : string = '') ;
Var StC,STitreCol :  String ;
    C  : integer ;
BEGIN
 // recuperation des nom des colonnes de la grille
StC:=G.Titres[0] ; C:=0 ;
Repeat
 STitreCol:=ReadTokenSt(StC) ; if STitreCol='' then exit ;
 if TOBLigne=nil then exit ;
 // on test si l'on doit utiliser les montants en monnaie de tenue, en contrevaleur ou en devise
 AnalyseTitreMontant(STitreCol,Devise) ;
 // si on ne force pas la colonne pour la refrence interne on recupere le parametrage de l'appli (param lettrage\lettrage gestion\refrence à afficher dans les traitements)
 if CRef='' then CRef:=VH^.CPRefLettrage ;
 // traitement du solde progressif
 if STitreCol='SOLDEPRO' then
  BEGIN
   if TOBLigne.GetValue('SOLDEPRO') < 0 then G.Cells[C,G.Row]:=StrS0(TOBLigne.GetValue('SOLDEPRO')*(-1)) + ' C'
   else if TOBLigne.GetValue('SOLDEPRO') > 0 then G.Cells[C,G.Row]:=StrS0(TOBLigne.GetValue('SOLDEPRO')) + ' D'
   else G.Cells[C,G.Row]:='0';
  END
 else
  BEGIN
   if (G.ColFormats[C]<>'') and ( VarType(TOBLigne.GetValue(STitreCol)) in [VarInteger,VarDouble,VarSingle] )
   then G.Cells[C,G.Row]:=StrS0(TOBLigne.GetValue(STitreCol))
   else G.Cells[C,G.Row]:=VarAsType(TOBLigne.GetValue(STitreCol),VarString) ;
   // traitement particulier de la reference interne
   if (STitreCol='E_REFINTERNE') then G.Cells[C,G.Row]:=TOBLigne.GetValue(GetChampRef(CRef));
  END;
 inc(C) ;
Until ((StC='') or (STitreCol='') or (C>=G.ColCount)) ;
// la dernier colonne non visible sert à forcer la selection de la grille
G.Cells[G.ColCount-1,G.RowCount-1]:=' ' ;
{$IFDEF TEST}
G.ColWidths[G.ColCount-1] := 13;
{$ENDIF}
END ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 16/11/2001
Modifié le ... : 03/11/2004
Description .. : Transforme un TOBM et TL_Rappro
Suite ........ : vIndex : index de OBM dans la piece. Permet de réaffecter
Suite ........ : l'OBM a partir du TL_Rappro d'origine
Suite ........ : - LG - 05/04/2002 - ajout des dates paquets max et min
Suite ........ : - LG - 03/11/2004 - FB 14765 - correction pour le lettrage 
Suite ........ : en devise
Mots clefs ... : 
*****************************************************************}
function COBMVersRappro( OBM : TOBM; vIndex : integer ; vBoLettrageDevise : boolean ) : TL_Rappro;
var
 L : TL_Rappro;
begin
 result:=nil; if OBM=nil then exit;
 L:=TL_Rappro.Create ;
 // Comptes et caractéristiques
 L.General:=OBM.GetMvt('E_GENERAL') ; L.Auxiliaire:=OBM.GetMvt('E_AUXILIAIRE') ;
 L.DateC:=OBM.GetMvt('E_DATECOMPTABLE') ; L.DateE:=OBM.GetMvt('E_DATEECHEANCE') ;
 L.DateR:=OBM.GetMvt('E_DATEREFEXTERNE'); L.RefI:=OBM.GetMvt('E_REFINTERNE') ;
 L.RefL:=OBM.GetMvt('E_REFLIBRE') ; L.RefE:=OBM.GetMvt('E_REFEXTERNE') ;
 L.Lib:=OBM.GetMvt('E_LIBELLE') ; L.Jal:=OBM.GetMvt('E_JOURNAL') ;
 L.Numero:=OBM.GetMvt('E_NUMEROPIECE'); L.NumLigne:=OBM.GetMvt('E_NUMLIGNE') ;
 L.NumEche:=OBM.GetMvt('E_NUMECHE') ; L.CodeL:=OBM.GetMvt('E_LETTRAGE') ;
 L.CodeD:=OBM.GetMvt('E_DEVISE') ; L.TauxDEV:=OBM.GetMvt('E_TAUXDEV') ;
 L.Nature:=OBM.GetMvt('E_NATUREPIECE') ;
 L.Facture:=((L.Nature='FC') or (L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF')) ;
 L.Client:=((L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC')) ;
 L.EditeEtatTva:=(OBM.GetMvt('E_EDITEETATTVA')='X') ;
 L.Solution:=0 ; L.Exo:=OBM.GetMvt('E_EXERCICE') ;
 L.EtatLettrage:=OBM.GetMvt('E_ETATLETTRAGE') ;
 // Montants, Lettrage
 L.DebDev:=OBM.GetMvt('E_DEBITDEV') ; L.CredDev:=OBM.GetMvt('E_CREDITDEV') ;
 L.Debit:=OBM.GetMvt('E_DEBIT') ; L.Credit:=OBM.GetMvt('E_CREDIT') ; L.Etat:=OBM.GetMvt('E_ETAT') ;
 L.DatePaquetMax := OBM.GetMvt('E_DATEPAQUETMAX') ; L.DatePaquetMin := OBM.GetMvt('E_DATEPAQUETMIN') ;
 if (L.CodeD<>V_PGI.DevisePivot) and vBoLettrageDevise then
  BEGIN
   L.DebitCur:=OBM.GetMvt('E_DEBITDEV') ; L.CreditCur:=OBM.GetMvt('E_CREDITDEV') ;
  END
   else
    BEGIN
     L.DebitCur:=OBM.GetMvt('E_DEBIT') ; L.CreditCur:=OBM.GetMvt('E_CREDIT') ;
    END;
 L.SoldePro:=OBM.GetMvt('SOLDEPRO');
 // on stocke l'index de OBM dans la piece pour retrouver OBM à partir du TL_Rappro
 L.NumLigneDep := vIndex; result := L;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/11/2001
Modifié le ... : 21/11/2001
Description .. : Transforme un OBM en TL_RAPPAUTO
Mots clefs ... :
*****************************************************************}
function COBMVersRappAuto( OBM : TOBM ) : T_RAPPAUTO;
var
 T : T_RAPPAUTO;
begin
 result:=nil; if OBM=nil then exit;
 T:=T_RAPPAUTO.Create ; T.General:=OBM.GetMvt('E_GENERAL') ;
 T.Auxiliaire:=OBM.GetMvt('E_AUXILIAIRE') ; T.Devise:=OBM.GetMvt('E_DEVISE') ;
 result := T ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 07/12/2001
Modifié le ... : 06/12/2002
Description .. : Chargement des options de lettrages
Suite ........ : LG : - ajout des nouveux param de lettrage
Suite ........ : -suppression des warnings
Suite ........ : 
Suite ........ : - 02/07/2002 - ajout d'une directive de compilation pour le 
Suite ........ : chargement des parametres
Suite ........ : -06/12/2002 - fiche 10763 - ajout de CAvecOD
Mots clefs ... : 
*****************************************************************}
function ChargeInfosLett : RINFOSLETT ;
Var
{$IFDEF SPEC302}
Q : TQuery ;
{$ENDIF}
IFL : RINFOSLETT;
BEGIN
FillChar(IFL,Sizeof(IFL),#0) ;
{$IFDEF SPEC302}
Q:=OpenSQL('Select * from SOCIETE',True,-1, '', True) ;
{références}
IFL.RefFC:=Q.FindField('SO_LETREFFC').AsString ; IFL.RefRC:=Q.FindField('SO_LETREFRC').AsString ;
IFL.RefFF:=Q.FindField('SO_LETREFFF').AsString ; IFL.RefRF:=Q.FindField('SO_LETREFRF').AsString ;
{Dates}
IFL.DateFC:=Q.FindField('SO_LETDATEFC').AsString ; IFL.DateRC:=Q.FindField('SO_LETDATERC').AsString ;
IFL.DateFF:=Q.FindField('SO_LETDATEFF').AsString ; IFL.DateRF:=Q.FindField('SO_LETDATERF').AsString ;
{Params}
IFL.EgalC:=(Q.FindField('SO_LETEGALC').AsString='X') ; IFL.EgalF:=(Q.FindField('SO_LETEGALF').AsString='X') ;
IFL.TolC:=Q.FindField('SO_LETTOLERC').AsInteger ; IFL.TolF:=Q.FindField('SO_LETTOLERF').AsInteger ;
Ferme(Q) ;
{$ELSE}
IFL.RefFC:=GetParamSoc('SO_LETREFFC') ; IFL.RefRC:=GetParamSoc('SO_LETREFRC') ;
IFL.RefFF:=GetParamSoc('SO_LETREFFF') ; IFL.RefRF:=GetParamSoc('SO_LETREFRF') ;
{Dates}
IFL.DateFC:=GetParamSoc('SO_LETDATEFC') ; IFL.DateRC:=GetParamSoc('SO_LETDATERC') ;
IFL.DateFF:=GetParamSoc('SO_LETDATEFF') ; IFL.DateRF:=GetParamSoc('SO_LETDATERF') ;
{Params}
// LG* nouveau param
IFL.ModeLettrage:=GetParamSoc('SO_LETMODE');
IFL.ChoixValid:=GetParamSoc('SO_LETCHOIXDEFVALID');
IFL.EcartDebit:=GetParamSoc('SO_LETECARTDEBIT');
IFL.EcartCredit:=GetParamSoc('SO_LETECARTCREDIT');
IFL.GeneralDebit:=GetParamSoc('SO_LETCHOIXGEN');
IFL.GeneralCredit:=GetParamSoc('SO_LETCHOIXGENC');
IFL.Journal:=GetParamSoc('SO_LETCHOIXJAL');
IFL.Libelle:=GetParamSoc('SO_LETCHOIXLIB');
IFL.AvecLetPartiel:=GetParamSoc('SO_LETMVTPARTIELLET');
IFL.LetTotal:=not GetParamSoc('SO_LETTOTAL');
IFL.AvecDebitCredit:=not GetParamSoc('SO_LETDC');
//IFL.MaxProf:=GetParamSoc('SO_LETMAXPROF');
//IFL.MaxDuree:=GetParamSoc('SO_LETMAXDUREE');
//IFL.JournalChange:=GetParamSoc('SO_LETJALCHANGE');
IFL.CAvecOD:=GetParamSoc('SO_LETMVTOD');
  {$IFDEF CCS3}
  IFL.EgalC:=False ; IFL.EgalF:=False ; IFL.TolC:=999 ; IFL.TolF:=999 ;
  {$ELSE}
  {Params}
  IFL.EgalC:=GetParamSoc('SO_LETEGALC') ; IFL.EgalF:=GetParamSoc('SO_LETEGALF') ;
  IFL.TolC:=GetParamSoc('SO_LETTOLERC') ; IFL.TolF:=GetParamSoc('SO_LETTOLERF') ;
  {$ENDIF}
{$ENDIF}
result:=IFL;
END ;

{=============================== RAPPRO REFERENCE ======================================}
function BonneDateRef ( LaDate,DateTest : TDateTime ; Client : boolean ; IFL : RINFOSLETT) : Boolean ;
BEGIN
if Client then BonneDateRef:=Abs(LaDate-DateTest)<=IFL.TolC
          else BonneDateRef:=Abs(LaDate-DateTest)<=IFL.TolF ;
END ;

function MemeRef ( LaRef,RefTest : String ; Client : boolean; IFL : RINFOSLETT ) : Boolean ;
Var Strict : boolean ;
BEGIN
if Client then Strict:=IFL.EgalC else Strict:=IFL.EgalF ;
if Strict then MemeRef:=(Laref=RefTest) else MemeRef:=((Pos(LaRef,RefTest)>0) or (Pos(RefTest,LaRef)>0)) ;
END ;

Function isISF ( L : TL_Rappro ; CODSALFOU : boolean) : Tisf ;
Var Q : TQuery ;
    Nat : String3 ;
BEGIN
Result:=isfDiv ;
if L.Auxiliaire<>'' then
   BEGIN
   Q:=OpenSQL('SELECT T_NATUREAUXI FROM TIERS WHERE T_AUXILIAIRE="'+L.Auxiliaire+'"',True,-1, '', True) ;
   if Not Q.EOF then Nat:=Q.Fields[0].AsString else Nat:='' ;
   if ((Nat='CLI') or (Nat='AUD')) then Result:=isfCli else
    if ((Nat='FOU') or (Nat='AUC')) then Result:=isfFou else
     if ((Nat='SAL') or (Nat='DIV')) then BEGIN if CODSALFOU then Result:=isfFou else Result:=isfCli ; END ;
   Ferme(Q) ;
   END else if L.General<>'' then
   BEGIN
   Q:=OpenSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="'+L.General+'"',True,-1, '', True) ;
   if Not Q.EOF then Nat:=Q.Fields[0].AsString else Nat:='' ;
   if Nat='TID' then Result:=isfCli else if Nat='TIC' then Result:=isfFou ;
   Ferme(Q) ;
   END ;
END ;

function QuelleRef ( L : TL_Rappro ; CAvecOD,CODSALFOU : boolean; IFL : RINFOSLETT) : String ;
Var Code : String3 ;
    St   : String ;
    isf : Tisf ;
BEGIN
if ((L.Nature='OD') and (CAvecOD)) then isf:=isISF(L,CODSALFOU) else isf:=isfDiv ;
if ((L.Nature='FC') or (L.Nature='RC') or (L.Nature='AC') or (L.Nature='OC') or (isf=isfCli)) then
   BEGIN
   if ((L.Facture) and (L.Nature<>'AC')) then Code:=IFL.RefFC else Code:=IFL.RefRC ;
   if isf=isfCli then
      BEGIN
      if ((L.Debit>0) or (L.Credit<0)) then Code:=IFL.RefFC else Code:=IFL.RefRC ;
      END ;
   END else if ((L.Nature<>'OD') or (isf=isfFou)) then
   BEGIN
   if ((L.Facture) and (L.Nature<>'AF')) then Code:=IFL.RefFF else Code:=IFL.RefRF ;
   if isf=isfFou then
      BEGIN
      if ((L.Credit>0) or (L.Debit<0)) then Code:=IFL.RefFF else Code:=IFL.RefRF ;
      END ;
   END ;
if Code='LIB' then St:=L.Lib else
 if Code='NUM' then St:=IntToStr(L.Numero) else
  if Code='RFE' then St:=L.RefE else
   if Code='RFI' then St:=L.RefI else
    if Code='RFL' then St:=L.RefL else St:='' ;
QuelleRef:=TripoteSt(St) ;
END ;

function QuelleDate ( L : TL_Rappro ; Var ControleDate : boolean ; CAvecOD,CODSALFOU : boolean;IFL : RINFOSLETT ) : TDateTime ;
Var Code : String3 ;
    DD   : TDateTime ;
    isf : Tisf ;
BEGIN
if ((L.Nature='OD') and (CAvecOD)) then isf:=isISF(L,CODSALFOU) else isf:=isfDiv ;
DD:=0 ; ControleDate:=True ;
if ((L.Nature='FC') or (L.Nature='RC') or (L.Nature='AC') or (L.Nature='OC')or (isf=isfCli)) then
   BEGIN
   if ((L.Facture) and (L.Nature<>'AC')) then Code:=IFL.DateFC else Code:=IFL.DateRC ;
   if isf=isfCli then
      BEGIN
      if ((L.Debit>0) or (L.Credit<0)) then Code:=IFL.DateFC else Code:=IFL.DateRC ;
      END ;
   END else if ((L.Nature<>'OD') or (isf=isfFou)) then
   BEGIN
   if ((L.Facture) and (L.Nature<>'AF')) then Code:=IFL.DateFF else Code:=IFL.DateRF ;
   if isf=isfFou then
      BEGIN
      if ((L.Credit>0) or (L.Debit<0)) then Code:=IFL.DateFF else Code:=IFL.DateRF ;
      END ;
   END ;
if Code='CPT' then DD:=L.DateC else
 if Code='ECH' then DD:=L.DateE else
  if Code='REF' then DD:=L.DateR else ControleDate:=False ;
Result:=DD ;
END ;

procedure AnnuleSolutionRef ( TLett : TList; NoSol : integer )  ;
Var i : integer ;
    X : TL_Rappro ;
BEGIN
for i:=0 to TLett.Count-1 do
    BEGIN
    X:=TL_Rappro(Tlett[i]) ; if X.Solution=NoSol then X.Solution:=0 ;
    END ;
END ;


Function BonneNature ( Nat1,Nat2 : String ; CAvecOD : boolean) : boolean ;
BEGIN
if (Nat1='FC') or (Nat1='AC') or (Nat1='RC') or (Nat1='OC') or ((Nat1='OD') and (CAvecOD))
   then Result:=(Nat2='FC') or (Nat2='AC') or (Nat2='RC') or (Nat2='OC') or ((Nat2='OD') and (CAvecOD))
   else if (Nat1='FF') or (Nat1='AF') or (Nat1='RF') or (Nat1='OF') or ((Nat1='OD') and (CAvecOD))
        then Result:=(Nat2='FF') or (Nat2='AF') or (Nat2='RF') or (Nat2='OF') or ((Nat2='OD') and (CAvecOD))
             else Result:=False ;
END ;


procedure DelettrePartiel ( X : TL_Rappro ) ;
Var SQL : String ;
BEGIN
SQL:='UPDATE ECRITURE SET E_LETTRAGE="", E_ETATLETTRAGE="AL", E_COUVERTURE=0, E_COUVERTUREDEV=0, ' ;
SQL:=SQL+'E_LETTRAGEDEV="-", E_REFLETTRAGE="", E_TRESOSYNCHRO = "LET" '; {JP 26/04/04 : pour l'échéancier de la Tréso}
SQL:=SQL+'Where E_AUXILIAIRE="'+X.Auxiliaire+'" AND E_GENERAL="'+X.General+'" ' ;
SQL:=SQL+'AND E_ETATLETTRAGE="PL" AND E_LETTRAGE="'+X.CodeL+'"' ;
ExecuteSQL(SQL) ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 08/04/2002
Modifié le ... : 02/05/2002
Description .. : LG - 08/04/2002 - Ajout du param CodeL qui contient le
Suite ........ : dernier code lettre ( sert au lettrage manuel )
Suite ........ :
Suite ........ : - LG - 02/05/2002 - ajout du param RL indiquant si on lettre
Suite ........ : en devise ( pour la gestion des ecarts de change )
Mots clefs ... :
*****************************************************************}
procedure FichierLettrage ( RL : RLETTR ; TLett : TList;NoSol : integer ; Total : boolean ; ModeRappro : Char ; var CodeL : string ; EcartPivot,EcartCV,EcartDev : double );
Var i : integer ;
    X : TL_Rappro ;
    Premier   : boolean ;
    LD,LC     : TList ;
    DMin,DMax : TDateTime ;
BEGIN
Premier:=True ; LD:=TList.Create ; LC:=TList.Create ;
DMin:=0 ; DMax:=0 ;
for i:=0 to TLett.Count-1 do
    BEGIN
    X:=TL_Rappro(TLett[i]) ; if X.Solution<>NoSol then Continue ;
    if ((Total) and (X.CodeL<>'')) then DelettrePartiel(X) ;
    if Premier then
       BEGIN
       CodeL:=GetSetCodeLettre(X.General,X.Auxiliaire) ; Premier:=False ;
       DMin:=X.DateC ; DMax:=X.DateC ;
       END else
       BEGIN
       if DMin>X.DateC then DMin:=X.DateC ;
       if DMax<X.DateC then DMax:=X.DateC ;
       END ;
    if Total then
       BEGIN
       X.EtatLettrage:='TL' ; X.CodeL:=uppercase(CodeL) ;
       if X.DebitCur<>0 then X.CouvCur:=X.DebitCur else X.CouvCur:=X.CreditCur ;
       END else
       BEGIN
       X.EtatLettrage:='PL' ; X.CodeL:=lowercase(CodeL) ;
       if X.DebitCur>0 then LD.Add(TLett[i]) else LC.Add(TLett[i]) ;
       END ;
       if not RL.LettrageDevise then
        BEGIN
         X.ConvertFranc:=EcartCV ; X.ConvertEuro:=EcartPivot ; X.EcartChangeFranc := EcartDev ; X.ModeRappro := ModeRappro ;
        END
         else
          BEGIN // ecart de change
           X.ConvertFranc:=0 ; X.ConvertEuro:=0 ; X.EcartRegulDevise:=0 ; X.ModeRappro := ModeRappro ;
           if VH^.TenueEuro then
            BEGIN
              X.EcartChangeFranc := EcartCV ; X.EcartChangeEuro:=EcartPivot ;
            END
             else
              BEGIN
               X.EcartChangeFranc := EcartPivot ; X.EcartChangeEuro:=EcartCV ;
              END;
          END;
          //  X.ModeRappro:=ModeRappro ;
    END ;
// réaffectation de la date paquet min et max
for i:=0 to TLett.Count-1 do
 BEGIN
  X:=TL_Rappro(TLett[i]) ; if X.Solution<>NoSol then Continue ;
  X.DatePaquetMax:=DMax ; X.DatePaquetMin:=DMin ;  X.ModeRappro:=ModeRappro ;
 END; // for
{$IFNDEF EAGLSERVER}
  if Not Total then RefEpuise(LD,LC) ;
{$ENDIF !EAGLSERVER}
LD.Clear ; LD.Free ; LC.Clear ; LC.Free ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 08/04/2002
Modifié le ... :   /  /
Description .. : LG - 08/04/2002 - Ajout du param CodeL qui contient le
Suite ........ : dernier code lettre ( sert au lettrage manuel )
Mots clefs ... :
*****************************************************************}
function SolutionsRef ( RL : RLETTR ; TLett : TList; NumSol,GeneNumSol : integer ; CAvecOD : boolean; var CodeL : string ; LetTotal : boolean=false ; AvecDebitCredit : boolean = false) : boolean;
Var NoSol,i : integer ;
    X       : TL_Rappro ;
    Nat    : String4 ;
    Debit,Credit : double ;
    DebitLettrage,CreditLettrage : double ;
    DebitDev,CreditDev : double ;
    Total,NatureCoherente,Premier : boolean ;
BEGIN
result:=false; if NumSol<=0 then Exit ; CodeL:='' ;
for NoSol:=1 to NumSol do
    BEGIN
    CodeL:='' ; Debit:=0 ; Credit:=0 ; NatureCoherente:=False ; Premier:=True ;
    DebitDev:=0 ; CreditDev:=0 ; DebitLettrage:=0 ; CreditLettrage:=0 ;
    for i:=0 to TLett.Count-1 do
        BEGIN
         X:=TL_Rappro(Tlett[i]) ; if X.Solution<>NoSol then Continue ;
         if Premier then
            BEGIN
            Nat:=X.Nature ; Premier:=False ;
            END else
            BEGIN
              // controle que les nature de piece ne sont pas identique sauf pour pour les type OD
              if X.Nature<>Nat then NatureCoherente:=True ;
              if ((X.Nature='OD') and (CAvecOD)) then NatureCoherente:=True ;
            END ;
         Debit:=Debit+X.Debit ; Credit:=Credit+X.Credit ;
         DebitLettrage:=DebitLettrage+X.DebitCur ; CreditLettrage:=CreditLettrage+X.CreditCur ;
         DebitDev:=DebitDev+X.DebDev ; CreditDev:=CreditDev+X.CredDev ;
         if ((CodeL='') and (X.CodeL<>'')) then CodeL:=X.CodeL ;
        END ; // for
    Total:=Arrondi(DebitLettrage-CreditLettrage,2)=0 ;
    result:=((CodeL='') or (Total)) ;
    // test si les ecritures lettres doivent avoir des valeurs au debit et au credit
    if AvecDebitCredit then
     result:= result and (Debit<>0) and (Credit<>0)
      else
       result:=((result) and ((Debit<>0) or (Credit<>0))) ;
    // test si on ne doit genere que du lettrage total
    if LetTotal then result:= result and Total;
    // calcul des ecart de conversion -> on se resert des variables sur le debit
    Debit     := Arrondi( Debit-Credit         , 2 ) ;
    DebitDev  := Arrondi( DebitDev-CreditDev   , 2 ) ;
    if not ToTal then BEGIN Debit:=0 ; DebitDev:=0 END ;
    Result    := result and NatureCoherente ;
    if result then FichierLettrage(RL,TLett,NoSol,Total,'R',CodeL,Debit,0,DebitDev) else AnnuleSolutionRef(TLett,NoSol) ;
    END ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/03/2002
Modifié le ... : 09/04/2002
Description .. : lettrage auto par rapport à la refrence
Suite ........ :
Suite ........ :
Suite ........ : LG - 08/04/2002 - Ajout du param CodeL qui contient le
Suite ........ : dernier code lettre ( sert au lettrage manuel )
Suite ........ : - on  ne prend pas en compte les ecritures partiellements
Suite ........ : lettrées
Mots clefs ... :
*****************************************************************}
function LettrerReference ( RL : RLETTR ; TLett : TList ; CAvecOD,CODSALFOU : boolean;IFL : RINFOSLETT;GeneNumSol : integer ; var CodeL : string) : boolean;
Var i,j,NumSol : integer ;
    L,X        : TL_Rappro ;
    LaRef,RefTest : String ;
    LaDate,DateTest : TDateTime ;
    Client,Okok,ControleDate,b,Premier : Boolean ;
BEGIN
NumSol:=0 ; CodeL:='' ;
for i:=0 to TLett.Count-1 do
    BEGIN
    L:=TL_Rappro(TLett[i]) ;
    if L.Solution>0 then Continue ; Premier:=True ;
    Client:=L.Client ;
    if Client then Okok:=(L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC') or ((L.Nature='OD') and (CAvecOD))
              else Okok:=(L.Nature='FF') or (L.Nature='AF') or (L.Nature='RF') or (L.Nature='OF') or ((L.Nature='OD') and (CAvecOD)) ;
    if Okok then
       BEGIN
       LaRef:=QuelleRef(L,CAvecOD,CODSALFOU,IFL) ; if LaRef='' then Continue ;
       LaDate:=QuelleDate(L,ControleDate,CAvecOD,CODSALFOU,IFL) ;
       for j:=0 to TLett.Count-1 do if j<>i then
           BEGIN
           X:=TL_Rappro(TLett[j]) ; if (X.Solution>0) or (X.EtatLettrage='PL') then Continue ;
           if Not BonneNature(L.Nature  ,X.Nature,CAvecOD) then Continue ;
           RefTest:=QuelleRef(X,CAvecOD,CODSALFOU,IFL) ; if RefTest='' then Continue ;
           if ControleDate then
              BEGIN
              DateTest:=QuelleDate(X,b,CAvecOD,CODSALFOU,IFL) ;
              if Not BonneDateRef(LaDate,DateTest,L.Client,IFL) then Continue ;
              END ;
           if MemeRef(LaRef,RefTest,L.Client,IFL) then
              BEGIN
              // pour chaque nouvelle solution on increment le numero de solution
              if Premier then BEGIN Inc(NumSol) ; Premier:=False ; END ;
              L.Solution:=NumSol ; X.Solution:=NumSol ;
              END ;
           END ;
       END ;
    END ;
 result:=SolutionsRef(RL,TLett,NumSol,GeneNumSol,CAvecOD,CodeL,IFL.letTotal,IFL.AvecDebitCredit) ;
END ;

function ConstruitListes ( TLett : TList ; icur : integer ; Debit : boolean ; Var LM : T_D ; Var LP,LG : T_I ) : integer ;
Var j,indice : integer ;
    X : TL_Rappro ;
BEGIN
FillChar(LM,Sizeof(LM),#0) ; FillChar(LP,Sizeof(LP),#0) ; FillChar(LG,Sizeof(LG),#0) ; Indice:=0 ;
for j:=0 to TLett.Count-1 do if j<>icur then
    BEGIN
    X:=TL_Rappro(TLett[j]) ;
    if (X.Solution=0) and (X.EtatLettrage='AL') then
       BEGIN
       if Debit then
          BEGIN
          if X.DebitCur<>0 then LM[Indice]:=-X.DebitCur else LM[Indice]:=X.CreditCur ;
          END else
          BEGIN
          if X.CreditCur<>0 then LM[Indice]:=-X.CreditCur else LM[Indice]:=X.DebitCur ;
          END ;
       LP[Indice]:=0 ; LG[Indice]:=j ; Inc(Indice) ;
       if Indice>=MaxDroite then Break ;
       END ;
    END ;
Result:=Indice ;
END ;

procedure CLettrerMontant(RL : RLETTR ; TLett : TList ; CAvecOD : boolean ; MaxProf,MaxDuree : integer ; Var CodeL : string) ;
Var i,Res,NumSol,NbD,k : integer ;
    L,X   : TL_Rappro ;
    LM    : T_D ;
    LP,LG : T_I ;
    Debit : boolean ;
    Solde : double ;
    Infos : REC_AUTO ;
    TT1   : TDateTime ;
    NbS   : Longint ;
    HH,MM,SS,CC : Word ;
    lRdDebit,lRdCredit         : double ;
    lRdDebitDev,lRdCreditDev   : double ;
BEGIN
TT1:=Time ; NumSol:=0 ;
for i:=0 to TLett.Count-1 do
    BEGIN
    L:=TL_Rappro(TLett[i]) ; 
    if (L.Solution<>0) or (L.EtatLettrage='PL') or (L.EtatLettrage='TL') then continue ;
    if (L.Facture) or ((L.Nature='OD') and CAvecOD) then
       BEGIN
       lRdDebit :=0 ; lRdCredit :=0 ; 
       lRdDebitDev := 0 ; lRdCreditDev := 0 ;
       Debit:=(L.DebitCur<>0) ; if Debit then Solde:=L.DebitCur else Solde:=L.CreditCur ;
       lRdDebit       := lRdDebit        + L.Debit ;
       lRdCredit      := lRdCredit       + L.Credit ;
       lRdDebitDev    := lRdDebitDev     + L.DebDev ;
       lRdCreditDev   := lRdCreditDev    + L.CredDev ;
       NbD:=ConstruitListes(TLett,i,Debit,LM,LP,LG) ;
       Infos.Nival:=MaxProf ; Infos.NbD:=NbD ; Infos.Decim:=2 ;
       Infos.Temps:=MaxDuree ; Infos.Unique:=False ;
       Res:=LettrageAuto(Solde,LM,LP,Infos) ;
       if Res=1 then
          BEGIN
          Inc(NumSol) ; L.Solution:=NumSol ;

          for k:=0 to NbD do if LP[k]<>0 then
           begin
            X              := TL_Rappro(TLett[LG[k]]) ;
            lRdDebit       := lRdDebit        + X.Debit ;
            lRdCredit      := lRdCredit       + X.Credit ;
            lRdDebitDev    := lRdDebitDev     + X.DebDev ;
            lRdCreditDev   := lRdCreditDev    + X.CredDev ;
            X.Solution     := NumSol;
           end; // for
           // calcul des ecart de conversion -> on se resert des variables sur le debit
           lRdDebit     := Arrondi( lRdDebit-lRdCredit         , 2 );
           lRdDebitDev  := Arrondi( lRdDebitDev-lRdCreditDev   , 2 );
           FichierLettrage(RL,TLett,NumSol,true,'C',CodeL,lRdDebit,0,lRdDebitDev) ;
          END ;
       END ;
    DecodeTime(Time-TT1,HH,MM,SS,CC) ;
    NbS:=3600*HH+60*MM+SS ; if NbS>MaxDuree then Break ;
    END ;
END ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/03/2002
Modifié le ... : 08/04/2002
Description .. : Lettrage auto par passage de solde progressif identique(voir
Suite ........ : jallaguier)
Suite ........ :
Suite ........ : LG - 08/04/2002 - Ajout du param CodeL qui contient le
Suite ........ : dernier code lettre ( sert au lettrage manuel )
Mots clefs ... :
*****************************************************************}
procedure LettrerSoldeProgressif(RL : RLETTR ; TLett : TList ; Var CodeL : string);
var
 i,j                        : integer;
 lInNumSol                  : integer;
 lRdSoldePro                : double;
 lRapproSuiv                : TL_Rappro;
 X                          : TL_Rappro;
 lRdDebit,lRdCredit         : double ;
 lRdDebitDev,lRdCreditDev   : double ;
 lRdDebitEuro,lRdCreditEuro : double ;
 lBoSolution                : boolean ;

 procedure Affect;
 var
  k : integer ;
 begin
   // on affecte la solution du point du depart a l'enregistrement courant
   for k := i to j do
    begin
     X              := TL_Rappro(TLett[k]) ;
     lRdDebit       := lRdDebit        + X.Debit ;
     lRdCredit      := lRdCredit       + X.Credit ;
     lRdDebitDev    := lRdDebitDev     + X.DebDev ;
     lRdCreditDev   := lRdCreditDev    + X.CredDev ;
     X.Solution     := lInNumSol;
    end; // for

   // calcul des ecart de conversion -> on se resert des variables sur le debit
   lRdDebit     := Arrondi( lRdDebit-lRdCredit         , 2 );
   lRdDebitDev  := Arrondi( lRdDebitDev-lRdCreditDev   , 2 );
   lRdDebitEuro := Arrondi( lRdDebitEuro-lRdCreditEuro , 2 );
   // on appel le lettrage pour cette solution
   FichierLettrage(RL,TLett,lInNumSol,True,'S',CodeL,lRdDebit,0,lRdDebitDev) ;
   // on passe au numero de solution suivant
   Inc(lInNumSol);
 end;

begin

 if TLett = nil then exit ;

 lInNumSol := 1 ; lRdSoldePro := 0 ; i := 0 ; lBoSolution := false ;

 while i < TLett.Count do
  begin

   lRdDebit :=0 ; lRdCredit :=0 ; lRdDebitEuro := 0 ; lRdCreditEuro := 0 ; lRdDebitDev := 0 ; lRdCreditDev := 0 ;
   // on part de l'enregistrement suivant pour rechercher un montant identique
   j := i + 1 ;
   // on va parcourir les enregistrements suivants pour trouver un solde progressif identique
   while ( not lBoSolution ) and ( j < TLett.Count ) do
    begin

     lRapproSuiv       := TL_Rappro( TLett[j] ) ;

     // il existe une solution ou l'ecriture est partiellement lettre -> on sort
     if ( lRapproSuiv.Solution > 0 ) or ( lRapproSuiv.EtatLettrage = 'PL' ) then break;

     if Arrondi(lRapproSuiv.SoldePro,2)=Arrondi(lRdSoldePro,2) then
      begin  // les soldes sont equivalent
       Affect ;
       // le nouveau solde progressif a rechercher est le denier solde du paquet lettre
       lRdSoldePro := TL_Rappro( TLett[j] ).SoldePro ;
       // on part de l'enregistrement suivant le dernier
       i           := j + 1;
       // on arrete la recherche
       lBoSolution := true ;
      end ; // if

     Inc(J) ;

    end; // while

   if not lBoSolution then
    begin
     // le solde progressif n'a pas ete trouvé -> on passe au suivant
     lRdSoldePro := TL_Rappro( TLett[i] ).SoldePro ;
     Inc(i) ;
    end;

   lBoSolution := false ;

  end; // while

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/03/2002
Modifié le ... : 26/04/2002
Description .. : Fct de lettrage auto par passage de solde à nulle
Suite ........ :
Suite ........ :
Suite ........ : LG - 08/04/2002 - Ajout du param CodeL qui contient le
Suite ........ : dernier code lettre ( sert au lettrage manuel )
Suite ........ :
Suite ........ : - LG - 23/04/2002 - on passe les ecritures totalement lettres
Suite ........ : de la grille
Suite ........ :
Suite ........ : - LG - 26/04/2002 - on calcule la valeur debit sur debit et
Suite ........ : credit et non plus sur debitCur et CreditCur
Mots clefs ... :
*****************************************************************}
procedure LettrerSolde( RL : RLETTR ; TLett : TList ; var CodeL : string) ;
Var i,k,NumSol,Depart : integer ;
    X : TL_Rappro ;
    L : TL_Rappro ;
    Solde : double ;
    PasDeSolution : boolean ;
    Debit,Credit : double ;
    DebitDev,CreditDev : double ;
BEGIN
NumSol:=0 ; Depart:=0 ;
Repeat
 Solde:=0 ; PasDeSolution:=True ; Debit :=0 ; Credit :=0 ; DebitDev := 0 ; CreditDev := 0 ;
 for i:=Depart to TLett.Count-1 do
     BEGIN
     X:=TL_Rappro(TLett[i]) ; if (X.Solution<>0) or (X.EtatLettrage='PL') or (X.EtatLettrage='TL') then break ;
     Solde:=Solde+X.DebitCur-X.CreditCur ;
     if Arrondi(Solde,2)=0 then
        BEGIN
        Inc(NumSol) ; PasDeSolution:=False ;
        for k:=Depart to i do if TL_Rappro(TLett[k]).Solution=0 then
         begin
          L:=TL_Rappro(TLett[k]); L.Solution:=NumSol ;
          Debit:=Debit+L.Debit ; Credit:=Credit+L.Credit ;
          DebitDev:=DebitDev+L.DebDev ; CreditDev:=CreditDev+L.CredDev ;
         end; // for
        Debit     := Arrondi( Debit-Credit         , 2 );
        DebitDev  := Arrondi( DebitDev-CreditDev   , 2 );
        FichierLettrage(RL,TLett,NumSol,True,'S',CodeL,Debit,0,DebitDev) ;
        Depart:=i+1 ; Break ;
        END ;
     END ;
Until ((Depart>=TLett.Count) or (PasDeSolution)) ;
END ;



function EstSelect ( G : THGrid ; Lig : integer ) : boolean ;
BEGIN
EstSelect:=Trim(G.Cells[G.ColCount-1,Lig])<>'' ;
END ;

Procedure ConvertCouverture ( O : TOBM ; Origine : TSorteMontant ) ;
Var Taux,XO : Double ;
    DEV : RDEVISE ;
BEGIN
if Origine=tsmPivot then
   BEGIN
    O.PutMvt('E_COUVERTUREDEV',O.GetMvt('E_COUVERTURE')) ;
   END ;
if Origine=tsmDevise then
   BEGIN
   XO:=O.GetMvt('E_COUVERTUREDEV') ; DEV.Code:=O.GetMvt('E_DEVISE') ; Taux:= O.GetMvt('E_TAUXDEV') ;
   GetInfosDevise(DEV) ; DEV.Taux:=Taux ;
   O.PutMvt('E_COUVERTURE',DeviseToEuro(XO,Taux,DEV.Quotite)) ; 
   END ;
END ;

function  TripoteSt( St : String ) : String ;
Var Res : String ;
    C   : Char ;
    i   : integer ;
BEGIN
st:=uppercase(Trim(St)) ; Res:='' ;
While Pos('  ',St)>0 do Delete(St,Pos('  ',St),1) ;
for i:=1 to length(St) do
    BEGIN
    C:=St[i] ;
    Case C of
       'é','è','ê','ë','É'     : St[i]:='E' ;
       'à','á','â','ä','å','Ä' : St[i]:='A' ;
       'í','ï','î'             : St[i]:='I' ;
       'ó','Ö','ô','ö','ò'     : St[i]:='O' ;
       'û','ù','Ü','ü'         : St[i]:='U' ;
       'ÿ'                     : St[i]:='Y' ;
       'Ç','ç'                 : St[i]:='C' ;
       END ;
    Res:=Res+St[i] ;
    END ;
Result:=Res ;
END ;

Procedure MajTabTvaEnc ( O : TOBM ; OkL : boolean ) ;
Var Nat,Jal,Regime,CodeTva : String3 ;
    i       : integer ;
    Client  : boolean ;
    Taux,TTC,HT : double ;
    Coll        : String ;
    Q : TQuery ;
BEGIN
{#TVAENC}
if Not VH^.OuiTvaEnc then Exit ;
Nat:=O.GetMvt('E_NATUREPIECE') ; if ((Nat='FC') or (Nat='AC') or (Nat='FF') or (Nat='AF')) then Exit ;
Jal:=O.GetMvt('E_JOURNAL') ; if EstJalFact(Jal) then Exit ;
If Nat = 'OD' then  {27/11/07 YMO On ne reinit pas les infos TVA pour les OD autres que Effet (ex : OD de compte à compte)}
begin
  Q:=OpenSQL('SELECT 1 FROM JOURNAL WHERE J_JOURNAL="'+Jal+'" AND J_EFFET<>"X"',True,-1, '', True) ;
  if not (Q.EOF) then begin ferme(Q); Exit; end;
end;
if O.GetMvt('E_EDITEETATTVA')='X' then Exit ;
if OkL then
   BEGIN
   {Lettrage --> si pas facture et pas etat tva alors mise à 0 du tableau}
   O.PutMvt('E_ECHEENC1',0) ;
   END else
   BEGIN
   {Dé-lettrage --> si pas facture et pas etat tva alors recalcul tableau sur taux directeur}
   Coll:=O.GetMvt('E_GENERAL') ;
   if EstCollFact(Coll) then
      BEGIN
      CodeTva:=VH^.NumCodeBase[1] ; if CodeTva='' then Exit ;
      Regime:=O.GetMvt('E_REGIMETVA') ; Client:=((Nat='FC') or (Nat='AC')) ;
      Taux:=Tva2Taux(Regime,CodeTva,Not Client) ;
      if Client then TTC:=O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT')
                else TTC:=O.GetMvt('E_CREDIT')-O.GetMvt('E_DEBIT') ;
      HT:=Arrondi(TTC/(1.0+Taux),V_PGI.OkDecV) ; O.PutMvt('E_ECHEENC1',HT) ;
      END else
      BEGIN
      O.PutMvt('E_ECHEENC1',0) ;
      END ;
   END ;
for i:=2 to 4 do O.PutMvt('E_ECHEENC'+IntToStr(i),0) ;
O.PutMvt('E_ECHEDEBIT',0) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 02/01/2002
Modifié le ... :   /  /    
Description .. : Suppression du RMVT
Mots clefs ... :
*****************************************************************}
function GoReqMajLet ( O : TOBM ; Gene,Auxi : String17 ; Stamp : TDateTime ; OKL : boolean ) : boolean ;
Var
S,SQL : String ;
DD    : TDateTime ;
BEGIN
//LG* 02/02/2002
if VH^.OuiTvaEnc then MajTabTvaEnc(O,OkL) ;
O.PutValue('E_PAQUETREVISION',1) ;
{JP 23/09/04 : Cela évite d'avoir à modifier StPourUpdate, comme c'était le cas lorsque j'affectais
               manuellement E_TRESOSYNCHRO}
// Pb CDL, compte lettrable mais e_tresosynchro doit rester à RIE... SBO 23/08/2005
//if O.GetString('E_TRESOSYNCHRO')<>'RIE' then // en attente car les autres modifs ne sont pas faciles...
                                               // Etudier la possibilité d'aghir au nivo de la synchro ?
O.PutValue('E_TRESOSYNCHRO', 'LET') ;
S:=O.StPourUpdate ; DD:=O.GetValue('E_DATEMODIF') ;
If Trim(S)<>'' Then
  BEGIN
  SQL:='UPDATE ECRITURE SET '+S+', E_DATEMODIF="'+UsTime(Stamp) + '"'
      +' Where E_GENERAL="'+Gene+'" AND E_AUXILIAIRE="'+Auxi+'"'
      +' AND E_JOURNAL="'+O.GetMvt('E_JOURNAL')+'" AND E_EXERCICE="'+O.GetMvt('E_EXERCICE')+'"'
      +' AND E_DATECOMPTABLE="'+UsDateTime(O.GetMvt('E_DATECOMPTABLE'))+'" AND E_NUMEROPIECE='+InttoStr(O.GetMvt('E_NUMEROPIECE'))
      +' AND E_QUALIFPIECE="'+O.GetMvt('E_QUALIFPIECE')+'" AND E_NATUREPIECE="'+O.GetMvt('E_NATUREPIECE')+'"'
      +' AND E_NUMLIGNE='+IntToStr(O.GetMvt('E_NUMLIGNE'))+' AND E_NUMECHE='+IntToStr(O.GetMvt('E_NUMECHE'))
      +' AND E_DATEMODIF="'+UsTime(DD)+'"' ;
  try
   Result:=(ExecuteSQL(SQL)=1) ;
   if not result then
     PGIInfo('La ligne suivante : '+#10#13+
     'Exercice : '+O.GetValue('E_EXERCICE')+' journal : '+ O.GetValue('E_JOURNAL')+#10#13+
     'Bordereau n° '+intToStr(O.GetValue('E_NUMEROPIECE'))+' du '+ DateToStr(O.GetValue('E_DATECOMPTABLE'))+' libellé '+ O.GetValue('E_LIBELLE')+
     ' débit : '+StrS0(O.GetValue('E_DEBIT'))+' crédit : '+StrS0(O.GetValue('E_CREDIT'))+#10#13+' n''a pas pu être mise à jour ! ','Erreur dans l''écriture');
  except
   on E:Exception do
    BEGIN
     result:= false ; MessageAlerte(E.Message) ;
    END;
  end;
  END Else Result:=TRUE ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Gendreau laurent
Créé le ...... : 30/06/2003
Modifié le ... : 30/06/2003
Description .. : - LG - 30/06/2003 - ajout de new param.
Suite ........ : LettrageEnSaisie : si true on ouvre le lettrage avec les
Suite ........ : ecritures de type L ( lettrage en saisie de la saisie bordereau
Suite ........ : )
Suite ........ : CritEtatLettrage : FB 12342 - chaine sql sur 
Suite ........ : E_ETATLETTRAGE provenant de la consultation des 
Suite ........ : ecritures. Permet de choisir quelles types d'ecritures 
Suite ........ : apparaisse ds la lettrage
Mots clefs ... :
*****************************************************************}
Function LWhereBase ( OkLettrage,OkTiers,SaisieTreso,SansLesPartiels : boolean ; LettrageEnSaisie : boolean = false ; CritEtatLettrage : string = '' ) : String ;
Var StConf,StV8 : String ;
BEGIN
if SaisieTreso then
   BEGIN
   if OkLettrage Then
      BEGIN
      if LettrageEnSaisie then
      Result:=' (( (E_QUALIFPIECE="N" OR E_QUALIFPIECE="L") AND E_TRESOLETTRE<>"X") OR (E_QUALIFPIECE="'+V_PGI.User+'")) '
             +'AND E_NUMECHE>0 AND E_ECHE="X" AND E_ETATLETTRAGE<>"RI" '
             +'AND ((E_DEBIT+E_CREDIT)<>0 OR (E_DEBITDEV+E_CREDITDEV)<>0)'
      else
      Result:=' ((E_QUALIFPIECE="N" AND E_TRESOLETTRE<>"X") OR (E_QUALIFPIECE="'+V_PGI.User+'")) '
             +'AND E_NUMECHE>0 AND E_ECHE="X" AND E_ETATLETTRAGE<>"RI" '
             +'AND ((E_DEBIT+E_CREDIT)<>0 OR (E_DEBITDEV+E_CREDITDEV)<>0)' ;
      if SansLesPartiels then Result:=Result+' AND E_ETATLETTRAGE="AL"' ;
      END else
      BEGIN
      if LettrageEnSaisie then
      Result:=' ((E_QUALIFPIECE="N" OR E_QUALIFPIECE="L") OR (E_QUALIFPIECE="'+V_PGI.User+'")) '
             +'AND E_NUMECHE>0 AND E_ECHE="X" AND E_ETATLETTRAGE<>"RI" '
             +'AND ((E_DEBIT+E_CREDIT)<>0 OR (E_DEBITDEV+E_CREDITDEV)<>0)'
      else
      Result:=' ((E_QUALIFPIECE="N") OR (E_QUALIFPIECE="'+V_PGI.User+'")) '
             +'AND E_NUMECHE>0 AND E_ECHE="X" AND E_ETATLETTRAGE<>"RI" '
             +'AND ((E_DEBIT+E_CREDIT)<>0 OR (E_DEBITDEV+E_CREDITDEV)<>0)' ;
      END ;
   END else
   BEGIN
   if LettrageEnSaisie then
   Result:=' ( E_QUALIFPIECE="N" OR E_QUALIFPIECE="L" ) AND E_TRESOLETTRE<>"X" AND E_NUMECHE>0 '
          +'AND E_ECHE="X" AND E_ETATLETTRAGE<>"RI" '
          +'AND ((E_DEBIT+E_CREDIT)<>0 OR (E_DEBITDEV+E_CREDITDEV)<>0)'
   else
   Result:=' E_QUALIFPIECE="N" AND E_TRESOLETTRE<>"X" AND E_NUMECHE>0 '
          +'AND E_ECHE="X" AND E_ETATLETTRAGE<>"RI" '
          +'AND ((E_DEBIT+E_CREDIT)<>0 OR (E_DEBITDEV+E_CREDITDEV)<>0)' ;
   END ;
if CritEtatLettrage <> '' then result:=result+CritEtatLettrage
 else
   if OkLettrage then Result:=Result+' AND E_ETATLETTRAGE<>"TL" ' else Result:=Result+' AND E_ETATLETTRAGE<>"AL" ' ;
Result:=Result+' AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") ' ;
{$IFNDEF EAGLCLIENT}
StConf:=SQLConf('ECRITURE') ;
{$ENDIF}
if StConf<>'' then Result:=Result+' AND '+StConf ;
StV8:=LWhereV8 ; if StV8<>'' then Result:=Result+' AND '+StV8 ;
END ;

Function LTri : String ;
BEGIN
LTri:=' order by E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE, E_LETTRAGE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ' ;
END ;

procedure RemplirInfosRegul ( Var J : REGJAL ; Value : String );
Var Q     : TQuery ;
    i     : integer ;
BEGIN
FillChar(J,Sizeof(J),#0) ; if Value='' then Exit ;
Q:=OpenSQL('Select J_JOURNAL, J_COMPTEURNORMAL, '
          +'J_CPTEREGULDEBIT1, J_CPTEREGULDEBIT2, J_CPTEREGULDEBIT3, '
          +'J_CPTEREGULCREDIT1, J_CPTEREGULCREDIT2, J_CPTEREGULCREDIT3, J_MODESAISIE from JOURNAL '
          +'Where J_JOURNAL="'+Value+'"',True,-1, '', True) ;
if Not Q.EOF then
   BEGIN
   J.Journal:=Value ; J.Facturier:=Q.FindField('J_COMPTEURNORMAL').AsString ;
   for i:=1 to 3 do
       BEGIN
       J.D[i]:=Q.FindField('J_CPTEREGULDEBIT'+InttoStr(i)).AsString ;
       J.C[i]:=Q.FindField('J_CPTEREGULCREDIT'+InttoStr(i)).AsString ;
       END ;
   J.ModeSaisie:=Q.FindField('J_MODESAISIE').AsString ;
   END ;
Ferme(Q) ;
END ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/04/2002
Modifié le ... : 29/09/2003
Description .. : Retourne l'ensemble des ligne selectionne de la grille
Suite ........ : 
Suite ........ : LG - 29/09/2003 - FB 12698 - on ne pouvait pas lance 
Suite ........ : plusieur annulation a la suite
Mots clefs ... : 
*****************************************************************}
function CGetTOBSelectFromGrid( G : THGrid ; vBoTotal : boolean=false) : TOBM;
var
 j     : integer;
 O1,O2 : TOBM;
 O : TOBM;
begin
 result := nil ;
 if (G=nil) then exit ; O:=TOBM.Create(EcrGen,'',false) ;
 for j:=0 to G.RowCount-1 do
  begin
  if vBoTotal or EstSelect(G,j) then
   begin
    O1:=GetO(G,j); if O1=nil then continue ;
    O2 := TOBM.Create ( EcrGen, '', False , O) ;
    EgaliseOBM (O1,O2);
   end; // if
  end; // for
 result:=O;
end;


procedure CAjouteChampsSuppLett ( O : TOB );
begin
 O.AddChampSup('PASENREGISTRER',False) ; O.AddChampSup('SELECTIONNER',False) ;
 O.AddChampSupValeur('SOLDEPRO', 0 ,False) ; O.AddChampSup('ENBASE',False) ;
end;

procedure CEtudieModeLettrage ( var R : RLETTR ) ;
var i : integer ;
begin
  R.LettrageDevise := False ;
  if R.DeviseMvt = V_PGI.DevisePivot then
  begin
    R.CritDev := V_PGI.DevisePivot ;
  end
  else
  begin
    if R.DeviseMvt=R.CritDev then
    begin
      {Paquet en devise --> Lettrage devise}
      R.LettrageDevise:=True ;
    end
    else
    begin
      {Paquet en devise (non explicite) --> poser la question}
      i := PgiAsk('Les mouvements sont en devise. Voulez-vous lettrer en devise ?', 'Lettrage manuel');
      R.LettrageDevise := (i = mrYes) ;
      if R.LettrageDevise then
        R.CritDev := R.DeviseMvt
      else
        R.Distinguer := False ;
    end ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/10/2003
Modifié le ... :   /  /    
Description .. : - 14/10/2003 - LG - le champs e_paquetrevision n'etait pas
Suite ........ : affecte
Mots clefs ... : 
*****************************************************************}
function _ConstruireRequeteDelettrage : string ;
begin
 result := 'update ecriture set e_datepaquetmax=e_datecomptable,e_datepaquetmin=e_datecomptable,e_lettrage="",'+
           'e_etatlettrage="AL",e_lettragedev="-",e_couverture=0,e_couverturedev=0, '+
           'e_reflettrage="",e_etat="0000000000", e_datemodif="'+UsTime(NowH)+'" , E_TRESOSYNCHRO = "LET" ';{JP 26/04/04 : pour l'échéancier de la Tréso}
result:=result+', E_PAQUETREVISION=1 ' ;
end;

procedure CExecDelettreMvtExoRef( vStGeneral, vStAux : string );
var
 lStSql : string;
begin
  lStSQL := _ConstruireRequeteDelettrage ;
  lStSQL := lStSQL + ' WHERE E_GENERAL="' + vStGeneral + '" AND E_AUXILIAIRE="' + vStAux + '" ' +
            ' AND E_EXERCICE="' + VH^.CPExoRef.Code + '" AND E_DATEPAQUETMIN >= "' + UsDateTime(VH^.CPExoRef.Deb)
            + '" AND E_DATEPAQUETMAX <= "' + UsDateTime(VH^.CPExoRef.Fin) + '" and E_LETTRAGE<>"" '  ;
  ExecuteSQL(lStSQL);
end;


procedure CExecDeLettrage( vStGeneral, vStAux : string );
begin
  ExecuteSQL(_ConstruireRequeteDelettrage +
             ' where e_general="'+vStGeneral+'" ' +
             'and e_auxiliaire="'+vStAux+'" and E_LETTRAGE<>"" ');
  ExecuteSQL('update generaux set g_dernlettrage="", G_DATEDERNMVT="'+UsTime(VH^.CPExoRef.Deb)+'",g_datemodif="'+UsTime(NowH)+'"'+ ' where g_general="'+vStGeneral+'" ' );
  // FQ 10434 - CA - 06/02/2003
  if vStAux <> '' then
    ExecuteSQL('update tiers set t_dernlettrage="", t_DATEDERNMVT="'+UsTime(VH^.CPExoRef.Deb)+'",t_datemodif="'+UsTime(NowH)+'"'+ ' where t_auxiliaire="'+vStAux+'" ' );
end;

{$IFDEF COMPTA}
procedure CDeLettrageMulti ( G : THGrid ; vStPrefixe : string = 'G_' ; vBoMvtSurExo : boolean = false ) ;
var
 i          : integer ;
 lTOB       : TOB ;
 lListeBor  : TOB ;
 lStGeneral : string ;
 lStAux     : string ;
 lStMessage : string ;
begin                           
 //LG* 05/01/2002
 lListeBor:=CGetListeBordereauBloque ;
 try
 try
 if (G.AllSelected) or (G.nbSelected <> 0) then
  begin
   HGBeginUpdate(G) ;  i:=1 ;
   if G.AllSelected then
    BEGIN
      lTob := Tob(G.Objects[0, i]);
      if ( vStPrefixe = 'E' ) and not CControleLigneBor(G,lListeBor) then exit ;
      while ( lTob=nil ) do
       begin
        lStGeneral:=lTOB.GetValue(vStPrefixe+'GENERAL') ;
        lStAux:=lTOB.GetValue(vStPrefixe+'AUXILIAIRE') ; if lStAux=#0 then lstAux:='' ;
        if vBoMvtSurExo then CExecDelettreMvtExoRef(lStGeneral,lStAux) else CExecDeLettrage(lStGeneral,lStAux) ;
        Inc(i) ; lTob:=Tob(G.Objects[0, i]) ;
       end;
    END
    else // Traitement <> si AllSelected car sinon ça merde, Si AllSelected alors NbSelected est faux, il vaut 0
    BEGIN
     for i:=0 to G.nbSelected -1 do
      begin
       G.GotoLeBookMark(i) ; lTob:=Tob(G.Objects[0, G.Row]) ;
       if ( vStPrefixe = 'E' ) and  not CControleLigneBor(G,lListeBor) then exit ;
       lStGeneral:=lTOB.GetValue(vStPrefixe+'GENERAL') ;
       lStAux:=lTOB.GetValue(vStPrefixe+'AUXILIAIRE') ; if lStAux=#0 then lstAux:='' ;
       if vBoMvtSurExo then CExecDelettreMvtExoRef(lStGeneral,lStAux) else CExecDeLettrage(lStGeneral,lStAux) ;
      end;
    END;
  end; //if
 finally
  lListeBor.Free ;
 end;
 except
 on E:Exception do
  begin
   if V_PGI.SAV then lStMessage := ' CDeLettrageMulti : ' ;
   lStMessage                   := lStMessage + 'Erreur lors du delettrage. Toutes les opérations ont été annulées !' +#10#13#10#13 + E.message ;
   MessageAlerte(lStMessage ) ;
  end; // on
 end; // try
end;
{$ENDIF}


procedure _Copy( vPSource , vPDestination : TL_Rappro ) ;
begin
  vPDestination.General          := vPSource.General ;
  vPDestination.Auxiliaire       := vPSource.Auxiliaire ;
  vPDestination.NumTraChq        := vPSource.NumTraChq ;
  vPDestination.ChampMarque      := vPSource.ChampMarque ;
  vPDestination.ValeurMarque     := vPSource.ValeurMarque ;
  vPDestination.DebitCur         := vPSource.DebitCur ;
  vPDestination.CreditCur        := vPSource.CreditCur  ;
  vPDestination.CouvCur          := vPSource.CouvCur ;
  vPDestination.Debit            := vPSource.Debit ;
  vPDestination.CouvCur          := vPSource.CouvCur ;
  vPDestination.Debit            := vPSource.Debit ;
  vPDestination.Credit           := vPSource.Credit ;
  vPDestination.Couv             := vPSource.Couv ;
 // = vPDestination.DedDev         := // vPSource.DebDev
  vPDestination.CredDev          := vPSource.CredDev  ;
  vPDestination.CouvDev          := vPSource.CouvDev ;
  vPDestination.DateC            := vPSource.DateC ;
  vPDestination.DateE            := vPSource.DateE ;
  vPDestination.DateR            := vPSource.DateR ;
  vPDestination.RefI             := vPSource.RefI ;
  vPDestination.RefE             := vPSource.RefE ;
  vPDestination.RefL             := vPSource.RefL ;
  vPDestination.Lib              := vPSource.Lib ;
  vPDestination.Numero           := vPSource.Numero ;
  vPDestination.Jal              := vPSource.Jal ;
  vPDestination.Nature           := vPSource.Nature ;
  vPDestination.CodeD            := vPSource.CodeD ;
  vPDestination.NumLigne         := vPSource.NumLigne ;
  vPDestination.NumEche          := vPSource.NumEche ;
  vPDestination.CodeL            := vPSource.CodeL ;
  vPDestination.TauxDev          := vPSource.TauxDev ;
  vPDestination.Solution         := vPSource.Solution ;
  vPDestination.Decim            := vPSource.Decim ;
  vPDestination.Facture          := vPSource.Facture ;
  vPDestination.Client           := vPSource.Client ;
 /// vPDestination.EditeEtaTva      :=             // vPSource.EditeEtaTva
  vPDestination.SoldePro         := vPSource.SoldePro ;
  vPDestination.NumLigneDep      := vPSource.NumLigneDep ;
  vPDestination.EtatLettrage     := vPSource.EtatLettrage ;
  vPDestination.Etablissement    := vPSource.Etablissement ;
  vPDestination.Etat             := vPSource.Etat ;
  vPDestination.EcartRegulDevise := vPSource.EcartRegulDevise ;
  vPDestination.EcartChangeEuro  := vPSource.EcartChangeEuro  ;
  vPDestination.EcartChangeFranc := vPSource.EcartChangeFranc ;
  vPDestination.ConvertEuro      := vPSource.ConvertEuro ;
  vPDestination.ConvertFranc     := vPSource.ConvertFranc ;
  vPDestination.ModeRappro       := vPSource.ModeRappro ;
  vPDestination.DatePaquetMax    := vPSource.DatePaquetMax ;
  vPDestination.DatePaquetMin    := vPSource.DatePaquetMin ;
  vPDestination.Exo              := vPSource.Exo ;
end;

procedure _FinirLettrage ( TLett : TL_Rappro ) ;
var
 lSQL : string ;
 Q    : TQuery ;
begin

 if TLett.EtatLettrage = '' then exit ;

 lSQL := 'UPDATE ECRITURE SET E_PAQUETREVISION=1,E_LETTRAGE="' + Copy(TLett.CodeL,1,4) + '", ' +
         'E_ETATLETTRAGE="'  + TLett.EtatLettrage +'",' +
         'E_DATEPAQUETMIN="' + UsDateTime(TLett.DatePaquetMin) + '", ' +
         'E_DATEPAQUETMAX="' + UsDateTime(TLett.DatePaquetMax) + '", ' +
         'E_COUVERTURE=' + StrFPoint(TLett.CouvCur) + ', ' +
         'E_COUVERTUREDEV=' + StrFPoint(TLett.CouvCur) +', ' +
         'E_LETTRAGEDEV="-" ,E_ETAT="' + TLett.Etat + '", E_TRESOSYNCHRO = "LET" '; {JP 26/04/04 : pour l'échéancier de la Tréso}

 {27/11/07 YMO On ne reinit pas les infos TVA pour les OD autres que Effet (ex : OD de compte à compte)}
 {En effet, on doit retrouver les infos sur ce genre d'OD qui remplacent en quelque sorte la facture une fois lettrées avec le règlement}
 If TLett.Nature = 'OD' then
   Q:=OpenSQL('SELECT 1 FROM JOURNAL WHERE J_JOURNAL="'+TLett.Jal+'" AND J_EFFET<>"X"',True,-1, '', True) ;

 if VH^.OuiTvaEnc and RazTvaEnc(TLett) and (Q.EOF) then
  lSQL := lSQL  +', E_ECHEENC1=0, E_ECHEENC2=0, E_ECHEENC3=0, E_ECHEENC4=0, E_ECHEDEBIT=0 ' ;

 If TLett.Nature = 'OD' then Ferme(Q) ;

 lSQL := lSQL + 'Where E_GENERAL="'+TLett.General+'" AND E_AUXILIAIRE="'+TLett.Auxiliaire+'" AND E_EXERCICE="'+TLett.Exo+'" '
              + 'AND E_JOURNAL="'+TLett.Jal+'" AND E_DATECOMPTABLE="'+USDATETIME(TLett.DateC)+'" AND E_NUMEROPIECE='+IntToStr(TLett.Numero)+' '
              + 'AND E_NUMLIGNE='+IntToStr(TLett.NumLigne)+' AND E_NUMECHE='+IntToStr(TLett.NumEche) ;

 if ExecuteSQL(lSQL)<>1 then begin V_PGI.IoError:=oeUnknown ; Exit ; end ;

end ;


procedure _CAjoutePieceACharger( vListe : TStringList ; vP : TL_Rappro ) ;
var
 lStCrit : string ;
begin

 lStCrit := vP.Jal + ';' +
            vP.Exo  + ';' +
            UsDateTime(vP.DateC) + ';' +
            IntToStr(vP.Numero) ;
//            IntToStr(vP.NumLigne) ;

 if vListe.IndexOf(lStCrit) = - 1 then
  vListe.AddObject(lStCrit,vP) ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 02/11/2004
Modifié le ... :   /  /
Description .. : - 02/11/2004 - LG - FB 14443 - eclatement des pieces
Suite ........ : d'achats
Mots clefs ... :
*****************************************************************}
procedure _EclateMvt( vBoVentilDebit : boolean ; vListe : TStringList ; TLett : TList ; vPCur : TL_Rappro ; var vIndC : integer ; var vRdSolde : double ; vInSol : integer ) ;
var
 lP : TL_Rappro ;
 i  : integer ;
begin

 for i := 0 to TLett.Count - 1 do
  begin

   lP := TL_Rappro(TLett[i] ) ;

   if vBoVentilDebit then
    begin
     if ( lP.Solution = vInSol ) and ( lP.DebitCur <> 0 )  then
      begin
       if lP.DebitCur < vRdSolde then
        begin
         vRdSolde := vRdSolde - lP.DebitCur ;
         _CAjoutePieceACharger(vListe,lP) ;
         continue ;
        end;
       lP :=TL_Rappro.Create ;
       _Copy(vPCur,lP) ;
       _CAjoutePieceACharger(vListe,lP) ;
       lP.DebitCur     := vPCur.DebitCur - vRdSolde ;
       vPCur.DebitCur  := vRdSolde ;
       lP.Solution     := 0 ;
       lP.TypeLigne    := 2 ;
       TLett.insert(vIndC+1,lP) ;
       exit ;
     end ;
    end
     else
      begin
       if ( lP.Solution = vInSol ) and ( lP.CreditCur <> 0 )  then
        begin
         if lP.CreditCur < vRdSolde then
          begin
           vRdSolde := vRdSolde - lP.CreditCur ;
           _CAjoutePieceACharger(vListe,lP) ;
           continue ;
          end;
         lP :=TL_Rappro.Create ;
         _Copy(vPCur,lP) ;
         _CAjoutePieceACharger(vListe,lP) ;
         lP.CreditCur    := vPCur.CreditCur - vRdSolde ;
         vPCur.CreditCur := vRdSolde ;
         lP.Solution     := 0 ;
         lP.TypeLigne    := 2 ;
         TLett.insert(vIndC+1,lP) ;
         exit ;
        end ;
      end ;

  end ; // for

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/04/2004
Modifié le ... :   /  /
Description .. : - LG - 27/04/2004 - qd on ne trouve pas de solution, on
Suite ........ : remets le pointeur a nil
Mots clefs ... :
*****************************************************************}
function _TrouverCreditSuivant( TLett : TList ; var vPCur : TL_Rappro ; var vIndC : integer ) : boolean ;
var
 i : integer ;
begin
 result := true ;
 vPCur  := nil ;
 vIndC  := 0 ;
 for i := 0 to TLett.Count - 1 do
  begin
   vPCur := TL_Rappro(TLett[i]) ;
   if ( vPCur.Solution = 0 ) and ( vPCur.CreditCur <> 0 ) then
    begin
     vIndC := i ;
     exit ;
    end ;
  end ;
 vPCur  := nil ;
 vIndC  := 0 ;
 result := false ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 02/11/2004
Modifié le ... :   /  /    
Description .. : - 02/11/2004 - LG - FB 14443 - eclatement des pieces 
Suite ........ : d'achats
Mots clefs ... : 
*****************************************************************}
function _TrouverDebitSuivant( TLett : TList ; var vPCur : TL_Rappro ; var vIndD : integer ) : boolean ;
var
 i : integer ;
begin
 result := true ;
 vPCur  := nil ;
 vIndD  := 0 ;
 for i := 0 to TLett.Count - 1 do
  begin
   vPCur := TL_Rappro(TLett[i]) ;
   if ( vPCur.Solution = 0 ) and ( vPCur.DebitCur <> 0 ) then
    begin
     vIndD := i ;
     exit ;
    end ;
  end ;
 vPCur  := nil ;
 vIndD  := 0 ;
 result := false ;
end;

procedure _AnnuleSolutionRef ( TLett : TList; NoSol : integer )  ;
Var i : integer ;
    X : TL_Rappro ;
BEGIN
for i:=0 to TLett.Count-1 do
    BEGIN
    X:=TL_Rappro(Tlett[i]) ; if (X.Solution=NoSol) and (X.TypeLigne <> 2) then X.Solution:=0 ;
    END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/04/2004
Modifié le ... : 02/11/2004
Description .. : - LG  - 27/04/2004 - si on arriva pas a  trouver de solution
Suite ........ : au credit , le result est faux
Suite ........ : - 02/11/2004 - LG - FB 14443 - eclatement des pieces
Suite ........ : d'achats
Mots clefs ... :
*****************************************************************}
function _AlgoLettAnc ( vBoVentilDebit : boolean ; vListe : TStringList ; RL : RLETTR ; TLett : TList ; vPCur : TL_Rappro ; var vNbSol : integer ; var vInd , vInSol : integer ; vRdSolde : double ; vPDur : TL_Rappro = nil) : boolean ;
var
 lRdSolde  : double ;
 lCodeL    : string ;
 lBoResult : boolean ;
 lInNull   : integer ;
 lBoTotal  : boolean ;
 lRdSeuil  : double ;
begin

 if vPCur = nil then
  begin
   if vBoVentilDebit then
    begin
     if _TrouverDebitSuivant(TLett,vPCur,vInd) then
      vRdSolde := vPCur.DebitCur ;
    end
     else
      if _TrouverCreditSuivant(TLett,vPCur,vInd) then
       vRdSolde := vPCur.CreditCur ;
  end ;
 if vPCur = nil then
  begin
   _AnnuleSolutionRef ( TLett , vInSol ) ;
   result := false ; // LG 27/04/2004
   exit ;
  end;

 vPCur.Solution := vInSol ;

 if ( vPDur = nil ) and vBoVentilDebit then
  _TrouverCreditSuivant(TLett,vPDur,lInNull )
   else
    if vPDur = nil then
     _TrouverDebitSuivant(TLett,vPDur,lInNull) ;

 if vPDur = nil then
  begin
   if vNbSol <> 0 then
     FichierLettrage ( RL , TLett , vInSol , false , 'A' , lCodeL , 0 , 0 ,0 )
      else
       vPCur.Solution := -1 ; // pas de solution au debit, on suppprime cette solution de lettrage
   result   := false ;
   exit ;
  end;

 vPDur.Solution   := vInSol ;
 if vBoVentilDebit then
  lRdSolde :=  vPDur.CreditCur - vRdSolde
   else
    lRdSolde :=  vPDur.DebitCur - vRdSolde ;
 if lRdSolde > 0 then
  begin
   if vBoVentilDebit then
    lBoResult := _TrouverDebitSuivant(TLett,vPCur,vInd)
     else
      lBoResult := _TrouverCreditSuivant(TLett,vPCur,vInd) ;
   if not lBoResult then
    begin
     // il n'y a plus de credit disponible pour couvrir le montant au debit faut annuler le lettrage
     FichierLettrage ( RL , TLett , vInSol , false , 'A' , lCodeL , 0 , 0 ,0 ) ;
     result := false ;
     exit ;
    end ;
   if vBoVentilDebit then
    vRdSolde := vRdSolde + vPCur.DebitCur
     else
      vRdSolde := vRdSolde + vPCur.CreditCur ;
   result    := _AlgoLettAnc(vBoVentilDebit,vListe,RL,TLett,vPCur,vNbSol,vInd,vInSol,vRdSolde,vPDur);
  end
   else
    begin

     lBoTotal := true ; 
     if vBoVentilDebit then
      lRdSeuil := GetParamSocSecur('SO_LETECARTCREDIT',0)*(-1)
       else
        lRdSeuil := GetParamSocSecur('SO_LETECARTDEBIT',0)*(-1) ;

     if GetParamSocSecur('SO_LETCHOIXJAL','') <> '' then
      begin
       if ( lRdSolde > lRdSeuil ) and ( lRdSolde < 0 ) then
        begin
         if vBoVentilDebit then
          vPDur.Credit := vPCur.Debit
           else
            vPCur.Credit := vPDur.Debit ;
         lRdSolde := 0 ;
         lBoTotal := false ;
        end ;  // if
      end ;

     if lRdSolde < 0 then
      begin
       if vBoVentilDebit then
        _EclateMvt(vBoVentilDebit,vListe,TLett,vPCur,vInd,vPDur.Credit,vInSol)
        else
         _EclateMvt(vBoVentilDebit,vListe,TLett,vPCur,vInd,vPDur.Debit,vInSol) ;
      end ;
      
     FichierLettrage ( RL , TLett , vInSol , lBoTotal , 'A' , lCodeL , 0 , 0 ,0 ) ;
     Inc(vInSol) ;
     Inc(vNbSol) ;
     vPDur := nil ;
     result := true ;
    end ;


end ;



function _MemeLigne ( vP1,vP2 : TL_Rappro ) : boolean ;
begin
 result := ( vP1.Jal      = vP2.Jal )                 and
           ( vP1.Exo      = vP2.Exo )                 and
           ( vP1.DateC    = vP2.DateC )               and
           ( vP1.Numero   = vP2.Numero )              ;
       //    ( vP.NumLigne = vTOB.GetValue('E_NUMLIGNE') ) ;
end;

function _MemeLigneNumero ( vTOB : TOB ;  vP : TL_Rappro ) : boolean ;
begin
 result := ( vP.Jal      = vTOB.GetValue('E_JOURNAL') )                  and
           ( vP.Exo      = vTOB.GetValue('E_EXERCICE') )                 and
           ( vP.DateC    = vTOB.GetValue('E_DATECOMPTABLE') )            and
           ( vP.Numero   = vTOB.GetValue('E_NUMEROPIECE') )              and
           ( vP.NumLigne = vTOB.GetValue('E_NUMLIGNE') ) ;
end;


function _ChargePiece ( vP : TL_Rappro ) : TOB ;
var
 lQ   : TQuery ;
begin

 result := TOB.Create('', nil , - 1 ) ;

 try

 lQ := OpenSql( 'SELECT * FROM ECRITURE '      +
                'WHERE E_EXERCICE="'           + vP.Exo + '" ' +
                'AND E_JOURNAL="'              + vP.Jal + '" ' +
                'AND E_DATECOMPTABLE="'        + USDateTime(vP.DateC) + '" ' +
                'AND E_NUMEROPIECE='           + intToStr(vP.Numero)  +
                ' AND E_QUALIFPIECE="N" '       +
                ' ORDER BY E_NUMEROPIECE' , true ,-1, '', True) ;

 result.LoadDetailDB('ECRITURE','','',lQ,true) ;

 finally
  Ferme(lQ) ;
 end ;


end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/11/2004
Modifié le ... :   /  /    
Description .. : - 02/11/2004 - LG - FB 14443 - eclatement des pieces
Suite ........ : d'achats
Mots clefs ... : 
*****************************************************************}
procedure _MAJLigne( vBoVentilDebit : boolean ; vRL : RLETTR ; vTOB : TOB ; vP : TL_Rappro ; vRDEV : RDEVISE ) ;
begin
{$IFDEF COMPTA}
 if ( vBoVentilDebit ) and ( vTOB.GetValue('E_CREDIT') = 0 ) then
  CSetMontants(vTOB,vP.DebitCur,0,vRDEV,true) ;

 if not ( vBoVentilDebit ) and ( vTOB.GetValue('E_DEBIT') = 0 ) then
  CSetMontants(vTOB,0,vP.CreditCur,vRDEV,true) ;

 CTLettVersOBM(vP,TOBM(vTOB),vRL,vP.DebitCur<>0 ,false) ;
{$ENDIF}
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/11/2004
Modifié le ... :   /  /    
Description .. : - 02/11/2004 - LG - FB 14443 - eclatement des pieces 
Suite ........ : d'achats
Mots clefs ... : 
*****************************************************************}
function _CreationLigne(vBoVentilDebit : boolean ; vRL : RLETTR ; vTOBPiece : TOB ; vP : TL_Rappro ; vRDEV : RDEVISE ) : boolean ;
{$IFDEF COMPTA}
var
 lTOBLigne : TOB ;
 i : integer ;
{$ENDIF}
begin

 result := false ;
{$IFDEF COMPTA}
 lTOBLigne := TOB.Create('ECRITURE',vTOBPiece,-1) ;

 for i := 0 to vTOBPiece.Detail.Count - 2 do
  if vP.NumLigne = vTOBPiece.Detail[i].GetValue('E_NUMLIGNE') then
   begin
    lTOBLigne.Dupliquer(vTOBPiece.Detail[i],true,true) ;
    CSupprimerInfoLettrage(lTOBLigne) ; // on enleve les info de lettrage de la ligne dupliquée
    CRemplirInfoLettrage(lTOBLigne) ;
    if vBoVentilDebit then
     CSetMontants(lTOBLigne,vP.DebitCur,0,vRDEV,true)
      else
       CSetMontants(lTOBLigne,0,vP.CreditCur,vRDEV,true) ;
    if vP.Solution <> -1 then
     CTLettVersOBM(vP,TOBM(lTOBLigne),vRL,vP.DebitCur<>0 ,false ) ;
    lTOBLigne.PutValue('E_NUMLIGNE',vTOBPiece.Detail.Count) ;
    exit ;
   end ; // if
{$ENDIF}
end ;

procedure _IncRef ( vB : T_RAPPAUTO ; vP : TL_Rappro ) ;
begin
 if vP.DebitCur <> 0 then
  begin
   Inc(vB.NbD) ;
   vB.TotalD := vB.TotalD + vP.DebitCur ;
  end
   else
    begin
     Inc(vB.NbC) ;
     vB.TotalC := vB.TotalC + vP.CreditCur ;
    end ;
 vB.CodeZ := vP.CodeL ;
 if vB.CodeA = '' then
  vB.CodeA := vP.CodeL ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/11/2004
Modifié le ... :   /  /    
Description .. : - 02/11/2004 - LG - FB 14443 - eclatement des pieces 
Suite ........ : d'achats
Mots clefs ... : 
*****************************************************************}
function _CreationEcrParAnc(vBoVentilDebit : boolean ; vListe : TStringList ; vRL : RLETTR ; TLett : TList ; vRDEV : RDEVISE ; vB : T_RAPPAUTO = nil ) : boolean ;
{$IFDEF COMPTA}
var
 i          : integer ;
 k          : integer ;
 lTOBPiece  : TOB ;
 lTOBLigne  : TOB ;
 lPInfo     : TL_Rappro ;
 lPCur      : TL_Rappro ;
 lInNumCur  : integer ;
 lInd       : integer ;
{$ENDIF}
begin

 result := true ;

{$IFDEF COMPTA}

 for i := 0 to vListe.Count - 1 do
  begin

   lPInfo    := TL_Rappro(vListe.Objects[i]) ;
   lTOBPiece := _ChargePiece(lPInfo) ; // la TOB est cree ds la fonction
   lInNumCur := lPInfo.Numero ;

   if not CDetruitAncienPiece(lTOBPiece) then begin result := false ; exit ; end ;

   lInd      := 0 ;

   while lInd < TLett.Count do
    begin

      lPCur := TL_Rappro(TLett[lInd]) ;
      if ( lPCur.Numero <> lInNumCur ) then
       begin
        Inc(lInd) ;
        Continue ;
       end ; // if

      // on est sur le meme numero
      if lPCur.TypeLigne <> 2 then
       begin

        for k := 0 to lTOBPiece.Detail.Count - 1 do
         begin
          lTOBLigne := lTOBPiece.Detail[k] ;
          if _MemeLigneNumero(lTOBLigne,lPCur) then
           begin
            _MAJLigne(vBoVentilDebit,vRL,lTOBLigne,lPCur,vRDEV) ;
            _IncRef (vB,lPCur) ;
            lPCur.Free ;
            TLett.Delete(lInd) ;
            break ;
           end ; // if
         end

       end
         else
          begin
           // il s'agit d'une nouvelle ligne
           _CreationLigne(vBoVentilDebit,vRL,lTOBPiece,lPCur,vRDEV) ;
           _IncRef (vB,lPCur) ;
           lPCur.Free ;
           TLett.Delete(lInd) ;
          end ;

    end ; // while

   CMAJSaisie(lTOBPiece) ;
   if lTOBPiece <> nil then lTOBPiece.Free ;

  end ; // for i

{$ENDIF}
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 02/11/2004
Modifié le ... :   /  /    
Description .. : - 02/11/2004 - LG - FB 14443 - eclatement des pieces 
Suite ........ : d'achats
Mots clefs ... : 
*****************************************************************}
function CLettrageParAnc( vBoVentilDebit : boolean ; RL : RLETTR ; TLett : TList ; RDEV : RDEVISE ; vB : T_RAPPAUTO = nil ) : boolean ;
var
 lPCur       : TL_Rappro ;
 lIndC       : integer ;
 lInSol      : integer ;
 lListe      : TStringList ;
 lInd        : integer ;
 lP          : TL_Rappro ;
 j           : integer ;
 lVar        : integer ;
begin

 result := true ;

 if ( TLett = nil ) or ( TLett.Count = 0  ) or ( TLett.Count = 1 ) then exit ;

 lIndC  := 0 ;
 lInSol := 1 ;
 lPCur  := nil ;
 lVar   := 0 ;
 lListe := TStringList.Create ;

 try

 while _AlgoLettAnc( vBoVentilDebit , lListe , RL ,TLett , lPCur ,lVar ,lIndC , lInSol , 0) do lVar := 0 ;

 lInd := 0 ;

  while lInd < TLett.Count do
   begin
    lP := TL_Rappro(TLett[lInd]) ;
    if lP.Solution = 0 then
     begin
      lP.Free ;
      TLett.Delete(lInd) ;
     end // if
       else
        begin
         for j := 0 to lListe.Count - 1 do
          if _MemeLigne(TL_Rappro(lListe.Objects[j]),lP ) then lP.Etat := 'X' ;
         Inc(lInd) ;
        end ;
   end ; // while

  lInd := 0 ;

  while lInd < TLett.Count do
   begin
    lP := TL_Rappro(TLett[lInd]) ;
     if lP.Etat <> 'X' then
      begin
       _FinirLettrage(lP) ;
       _IncRef(vB,lP) ;
       lP.Free ;
       TLett.Delete(lInd) ;
      end
        else
         Inc(lInd);
   end ; // while

 // on lettre les pieces que l'on doit charger
  _CreationEcrParAnc(vBoVentilDebit,lListe,RL,TLett,RDEV,vB) ;

  result := true ;

 finally
  lListe.Free ;
 end;

end ;


function NextLigneLettre ( L : TList ; FromLig : integer ; Var Reste : double ;
                           ValTest : double ; Signe : TSigne ) : integer ;
Var i : integer ;
    okok : boolean ;
    M : Double ;
    X    : TL_Rappro ;
BEGIN
Result:=-1 ;
for i:=FromLig to L.Count-1 do
    BEGIN
    X:=TL_Rappro(L[i]) ;
    if X.DebitCur<>0 then M:=X.DebitCur else M:=X.CreditCur ;
    Reste:=M-X.CouvCur ; Okok:=BonSigne(Reste,ValTest,Signe) ;
    if ((Okok) and (Reste<>0)) then BEGIN Result:=i ; Break ; END ;
    END ;
END ;

function TrouveLesQuels ( LD,LC : TList ; Var L1,L2 : TList ; Var i1,i2 : integer ) : boolean ;
Var R1,R2 : double ;
BEGIN
L1:=Nil ; L2:=Nil ; Result:=False ;
i1:=NextLigneLettre(LD,0,R1,0,sgNeutre) ;
if i1>=0 then
   BEGIN
   L1:=LD ; i2:=NextLigneLettre(LC,0,R2,R1,sgEgal) ;
   if i2>=0 then L2:=LC else
      BEGIN
      i2:=NextLigneLettre(LD,i1+1,R2,R1,sgOppose) ; if i2>=0 then L2:=LD else Exit ;
      END ;
   END else
   BEGIN
   i1:=NextLigneLettre(LC,0,R1,0,sgNeutre) ; if i1<0 then Exit ;
// GP le 03/04/2000 vu  avec JLD
// L1:=LD ;
   L1:=LC ;i2:=NextLigneLettre(LC,i1+1,R2,R1,sgOppose) ; if i2>=0 then L2:=LC else Exit ;
  END ;
Result:=True ;
END ;

procedure RapprocheLesLignes ( L1,L2 : TList ; i1,i2 : integer ) ;
Var M1,C1,R1,M2,C2,R2 : Double ;
    X1,X2             : TL_Rappro ;
BEGIN
X1:=TL_Rappro(L1[i1]) ; if X1.DebitCur<>0 then M1:=X1.DebitCur else M1:=X1.CreditCur ; C1:=X1.CouvCur ; R1:=M1-C1 ;
X2:=TL_Rappro(L2[i2]) ; if X2.DebitCur<>0 then M2:=X2.DebitCur else M2:=X2.CreditCur ; C2:=X2.CouvCur ; R2:=M2-C2 ;
if BonSigne(R1,R2,sgEgal) then
   BEGIN
   if Abs(R1)>Abs(R2) then BEGIN R1:=R1-R2 ; R2:=0 ; END else BEGIN R2:=R2-R1 ; R1:=0 ; END ;
   END else
   BEGIN
   if Abs(R1)>Abs(R2) then BEGIN R1:=R1+R2 ; R2:=0 ; END else BEGIN R2:=R1+R2 ; R1:=0 ; END ;
   END ;
C1:=M1-R1 ; X1.CouvCur:=C1 ;
C2:=M2-R2 ; X2.CouvCur:=C2 ;
END ;

procedure RefEpuise ( LD,LC : TList ) ;
Var ii    : integer ;
    L1,L2 : TList ;
    i1,i2 : integer ;
    Okok  : boolean ;
BEGIN
ii:=0 ;
Repeat
 Okok:=TrouveLesquels(LD,LC,L1,L2,i1,i2) ;
 if Okok then RapprocheLesLignes(L1,L2,i1,i2) ;
 inc(ii) ;
Until ((Not Okok) or (ii>500)) ;
END ;


function GetSetCodeLettre ( vStGen , vStAux : string ; vStDernLettrage : string = '') : string ;
var
 lQ          : TQuery ;
 lStOldLet   : string ;
 lStNewLet   : String ;
 lBoQEOF     : boolean ;
begin

 lStNewLet := '' ;

 if vStAux <> '' then
  begin
   lQ        := OpenSQL('Select T_DERNLETTRAGE from TIERS Where T_AUXILIAIRE="' + vStAux + '"' ,true,-1, '', True) ;
   lStNewLet := lQ.FindField('T_DERNLETTRAGE').AsString ;
  end
   else
    begin
     lQ        := OpenSQL('Select G_DERNLETTRAGE from GENERAUX Where G_GENERAL="' + vStGen + '"',true,-1, '', True) ;
     lStNewLet := lQ.FindField('G_DERNLETTRAGE').AsString ;
    end ;

 lBoQEOF := lQ.EOF ;
 Ferme(lQ) ;

 try

 if not lBoQEOF then
  begin

   if vStDernLettrage='' then
    lStOldLet := lStNewLet
     else
      lStOldLet := vStDernLettrage ;

   if lStNewLet = lStOldLet then
    begin
      vStDernLettrage := CodeSuivant(lStNewLet) ;

      if vStAux <> '' then
       ExecuteSQL('update TIERS set T_DERNLETTRAGE="' + vStDernLettrage + '" Where T_AUXILIAIRE="' + vStAux +'"')
        else
         ExecuteSQL('update GENERAUX set G_DERNLETTRAGE="' + vStDernLettrage + '" Where G_GENERAL="' + vStGen + '"')  ;
    end
     else
      V_PGI.IoError := oeUnknown ;
  end
   else
    V_PGI.IoError := oeUnknown ;

 result := vStDernLettrage ;

except
 on E:Exception do
  begin
   MessageAlerte('Problème lors de la recherche du code lettre'+#10#13+E.message) ;
   raise ;
  end;
end;

end ;


{function  LettrageTotalTOB ( vTOBEcr : TOB ) : boolean ;
var
 i          : integer ;
 lTOB       : TOB ;
 lDMin      : TDateTime ;
 lDMax      : TDateTime ;
 lBoPremier : boolean ;
 lStCodeL   : string ;
begin

 result := false ;

 lBoPremier   := true ;
 lDMin        := iDate1900 ;
 lDMax        := iDate1900 ;

 if vTOBEcr.Detail.Count < 2 then exit ;

 for i := 0 to vTOBEcr.Detail.Count - 1 do
  begin

   lTOB := vTOBEcr.Detail[i] ;

   if lBoPremier then
    begin
     lStCodeL     := GetSetCodeLettre( lTOB.GetString('E_GENERAL') , lTOB.GetString('E_AUXILIAIRE')) ;
     lBoPremier   := false ;
     lDMin        := lTOB.GetValue('E_DATECOMPTABLE') ;
     lDMax        := lTOB.GetValue('E_DATECOMPTABLE') ;
    end
     else
      begin
       if lDMin > lTOB.GetValue('E_DATECOMPTABLE') then
        lDMin := lTOB.GetValue('E_DATECOMPTABLE') ;
       if lDMax < lTOB.GetValue('E_DATECOMPTABLE') then
        lDMax := lTOB.GetValue('E_DATECOMPTABLE') ;
      end ; // if

  end ; // for

 for i := 0 to vTOBEcr.Detail.Count - 1 do
  begin
   lTOB := vTOBEcr.Detail[i] ;
   lTOB.PutValue('E_DATEPAQUETMIN'     ,  lDMin ) ;
   lTOB.PutValue('E_DATEPAQUETMAX'     ,  lDMax) ;
   lTOB.PutValue('E_ETATLETTRAGE'      , 'TL' ) ;
   lTOB.PutValue('E_LETTRAGE'          , lStCodeL) ;
   lTOB.PutValue('E_COUVERTURE'        , lTOB.GetValue('E_DEBIT')+lTOB.GetValue('E_CREDIT')) ;
   lTOB.PutValue('E_COUVERTUREDEV'     , lTOB.GetValue('E_DEBITDEV')+lTOB.GetValue('E_CREDITDEV')) ;
   GoReqMajLet(TOBM(lTOB),lTOB.GetString('E_GENERAL'),lTOB.GetString('E_AUXILIAIRE'),Now,true) ;
  end ; // for

end ; }
{$ENDIF EAGLSERVER}

function CodeSuivant ( CL : String ) : String ;
Var i : integer ;
BEGIN
CL:=uppercase(Trim(CL)) ;
if Length(CL)<4 then CL:='AAAA' else
   BEGIN
   i:=4 ; While CL[i]='Z' do BEGIN CL[i]:='A' ; Dec(i) ; END ;
   CL[i]:=Succ(CL[i]) ;
   END ;
CodeSuivant:=Copy(CL,1,4) ;
END ;

{$IFNDEF EAGLSERVER}
{$IFDEF COMPTA}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 23/07/2007
Modifié le ... :   /  /    
Description .. : - LG - 23/07/2007 -  FB 21120 - controle sur les lignes 
Suite ........ : totalement lettres
Mots clefs ... : 
*****************************************************************}
procedure  CLettrageTOB ( vTOBEcr : TOB ) ;
var
 i        : integer ;
 lTOB     : TOB ;
 lLett    : TList ;
 lTL      : TL_Rappro ;
 lRL      : RLETTR ;
 lStCodeL : string ;
 lRdSolde : double ;
begin

 if ( vTOBEcr = nil ) or ( vTOBEcr.Detail = nil ) then exit ;

 lLett             := TList.Create;
 lTOB              := vTOBEcr.Detail[0] ;
 lRL.General       := lTOB.GetValue('E_GENERAL') ;
 lRL.Auxiliaire    := lTOB.GetValue('E_AUXILIAIRE') ;
 lRL.Appel         := tlMenu;
 lRL.CritDev       := lTOB.GetValue('E_DEVISE') ;
 lRL.DeviseMvt     := lTOB.GetValue('E_DEVISE');
 lRL.GL            := nil;
 lRL.CritMvt       := '' ;
 lRdSolde          := 0 ;

 try

 for i := 0 to vTOBEcr.Detail.Count - 1 do
  begin
   lTOB           := vTOBEcr.Detail[i] ;
   if lTOB.GetString('E_ETATLETTRAGE') = 'TL' then
    begin
     PGIError('La ligne est déjà lettrée !' + #10#13 +
              'Journal :' + lTOB.GetString('E_JOURNAL') + '  Pièce :' + lTOB.GetString('E_NUMEROPIECE')  +
              '  Ligne :' + lTOB.GetString('E_NUMLIGNE') + '  Lettrage :' + lTOB.GetString('E_LETTRAGE') ) ;
     exit ;
    end ;
   CAjouteChampsSuppLett (lTOB);
   lTL            := COBMVersRappro(TOBM(lTOB),i,false); // transforme OBM en TL_Rappro
   lTL.ModeRappro := 'A' ;
   lTL.Solution   := 1 ;
   lRdSolde       := lRdSolde + Arrondi(lTOB.GetValue('E_DEBIT') - lTOB.GetValue('E_CREDIT') , 2 ) ; 
   if lTL <> nil then lLett.Add(lTL) ;
  end ; // for

 if not GetParamSocSecur('SO_LETTOTAL',False) and ( lRdSolde <> 0 ) then
  exit ;

 FichierLettrage ( lRL , lLett , 1 , lRdSolde = 0 ,'*', lStCodeL , 0 , 0 ,0 );

 for i:=0 to lLett.Count-1 do
  begin
   lTL  := TL_Rappro(lLett[i]) ;
   lTOB := vTOBEcr.Detail[lTL.NumLigneDep] ;
   CTLettVersOBM( lTL, TOBM(lTOB) , lRL , true) ;
   CTLettVersOBM( lTL, TOBM(lTOB) , lRL , false) ;
  end; // for

 for i := 0 to vTOBEcr.Detail.Count - 1 do
  if not GoReqMajLet( TOBM(vTOBEcr.Detail[i]) , lRL.General , lRL.Auxiliaire, Now , true) then
   begin
    V_PGI.IoError := oeLettrage ;
    exit ;
   end ;

 finally
  LLett.Free ;
 end ;

end ;
{$ENDIF}
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 10/03/2006
Modifié le ... :   /  /
Description .. : - LG - 10/03/2006 - suprression d'uen fuite memoire
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
function CControlePresenceLettrage ( vTOB : TOB ) : boolean ;
var
 i          : integer ;
 lLett      : HTStringList ;
 lSql       : string ;
 lQ         : TQuery ;
 lStGen     : string ;
 lStAux     : string ;
 lInNbMem,lInNbBase : integer ;
begin

 result     := false ;
 if ( vTOB.Detail =  nil ) or ( vTOB.Detail.Count =  0 ) then exit ;  
 lInNbMem   := 0 ;
 lInNbBase  := 0 ;
 lLett      := HTStringList.Create ;
 lStGen     := vTOB.Detail[0].Getvalue('E_GENERAL') ;
 lStAux     := vTOB.Detail[0].Getvalue('E_AUXILIAIRE') ;

 try

 for i := 0 to vTOB.Detail.Count - 1 do
  if length(trim(vTOB.Detail[i].Getvalue('E_LETTRAGE')))>0 then
   begin
    if (lLett.IndexOf( vTOB.Detail[i].Getvalue('E_LETTRAGE') ) = -1)  then
     lLett.add(vTOB.Detail[i].Getvalue('E_LETTRAGE')) ;
    inc(lInNbMem) ;
   end ; // if

 for i:=0 to lLett.Count-1 do
  begin
   lSql:='select count(*) N from ECRITURE where E_GENERAL="'+lStGen+'" AND E_AUXILIAIRE="'+lStAux+'"' ;
   lSql:=lSql+' and E_LETTRAGE="'+lLett[i]+'" ' ;
   lQ:=OpenSql(lSql,true,-1, '', True) ;
   // on compte les mouvements partiellment lettre en base
   lInNbBase:=lInNbBase+lQ.findField('N').asInteger ;
   ferme(lQ) ;
 end;

 finally
  lLett.Free ;
 end ;

 result := lInNbBase = lInNbMem ;

end ;
{$ENDIF}

procedure DelettreEcritureinMem (TOBECR : TOB);
	procedure Delettre (TOBL : TOB);
  begin
    TOBL.PutValue('E_COUVERTURE',0);
    TOBL.PutValue('E_COUVERTUREDEV',0);
    TOBL.PutValue('E_DATEPAQUETMAX',iDate1900);
    TOBL.PutValue('E_DATEPAQUETMIN',iDate1900);
    TOBL.PutValue('E_LETTRAGE','');
    TOBL.PutValue('E_DATEMODIF',NowH);
  end;

var indice : Integer;
		TOBL : TOB;
begin
  for Indice := 0 to TOBECR.Detail.Count -1 do
  begin
    TOBL := TOBECR.detail[Indice];
    if (TOBL.GetString('E_ECHE')='X') and (TOBL.GetString('E_LETTRAGE')<>'') then
    begin
			TOBL := TOBECR.detail[Indice]; Delettre(TOBL);
    end;
  end;
end;

procedure DelettreEcriture (var MM : RMVT ; TOBECR : TOB);

	procedure DelettreUnReglement (TOBL : TOB);
  var StSql : string;
  begin
    StSql := 'UPDATE ECRITURE '+
  					 'SET E_COUVERTURE=0, E_COUVERTUREDEV=0, E_DATEPAQUETMAX="'+USDatetime(Idate1900)+'",'+
    				 'E_DATEPAQUETMIN="'+USDateTime(IDate1900)+'", E_ETATLETTRAGE="AL", E_LETTRAGE="",'+
             'E_DATEMODIF="'+UsDateTime(NowH)+'" '+
    				 'WHERE E_GENERAL="'+TOBL.GetString('E_GENERAL')+'" '+
    				 'AND E_AUXILIAIRE="'+TOBL.GetString('E_AUXILIAIRE')+'" '+
             'AND E_LETTRAGE="'+TOBL.GetString('E_LETTRAGE')+'" '+
             'AND  E_QUALIFPIECE="N" AND E_TRESOLETTRE<>"X" AND E_NUMECHE>0 '+
             'AND E_ECHE="X" AND E_ETATLETTRAGE<>"RI" AND ((E_DEBIT+E_CREDIT)<>0 OR (E_DEBITDEV+E_CREDITDEV)<>0) '+
             'AND E_ETATLETTRAGE<>"AL"  AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") '+
             'AND E_NATUREPIECE IN ("OC","OF","RC","RF","OD") ';
    ExecuteSql(StSql);
  end;

var Indice : Integer;
		TOBL : TOB;
    first : boolean;
begin
  first := True;
  for Indice := 0 to TOBECR.Detail.Count -1 do
  begin
		TOBL := TOBECR.detail[Indice];
    if (TOBL.GetString('E_ECHE')='X') and (TOBL.GetString('E_LETTRAGE')<>'') then
    begin
      if first then
      begin
        MM.CodeLettrage := TOBL.GetString('E_LETTRAGE');
      end;
    	DelettreUnReglement (TOBL);
    end;
  end;
end;

end.

