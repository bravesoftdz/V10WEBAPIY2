{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/02/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit RappType ;
interface

uses Classes,Hent1 ;

type TInfoMvt = Class
                JOURNAL            : String;
                EXERCICE           : String;
                DATECOMPTABLE      : TDateTime;
                NUMEROPIECE        : Integer;
                NUMLIGNE           : Integer;
                NUMECHE            : Integer;
                NUMVENTIL          : Integer;
                NATUREPIECE        : String ;
                AUXILIAIRE         : String;
                GENERAL            : String;
                DEVISE             : String ;
                AXE                : String;
                AXEPLUS            : Array[1..5] of String[2] ;
                SECTION            : String;
                MODEPAIE           : String3;
                DEBIT              : Double;
                CREDIT             : Double;
                DEBITMIN           : Double;
                DEBITMAX           : Double;
                CREDITMIN          : Double;
                CREDITMAX          : Double;
                DEBITMINEURO       : Double;
                DEBITMAXEURO       : Double;
                CREDITMINEURO      : Double;
                CREDITMAXEURO      : Double;
                COUVERTURE         : Double;
                DEBITDEV           : Double;
                CREDITDEV          : Double;
                COUVERTUREDEV      : Double;
                TAUXDEV            : Double;
                ECRANOUVEAU        : String;
                CONFIDENTIEL       : String ;
                LETTRAGE           : String;
                ETATLETTRAGE       : String;
                LETTRAGEDEV        : String;
                ECHE               : String;
                ANA                : String;
                DATEECHEANCE       : TDateTime;
                TYPEANALYTIQUE     : String ;
                REFINTERNE         : String ;
                ETABLISSEMENT      : String3;
                QUALIFPIECE        : String ;
                DEBITEURO          : Double;
                CREDITEURO         : Double;
                CHRONO             : Integer ;
                CHAMP              : String ;
                MODIF              : Boolean ;
                INDICETABLE        : Byte;
                MAXCONTROLE        : String ;
                MODESAISIE         : String ;
                End ;
type TInfoMvtBud = Class
                BUDJAL            : String;
                EXERCICE           : String;
                DATECOMPTABLE      : TDateTime;
                NUMEROPIECE        : Integer;
                NATUREBUD          : String ;
                BUDGENE            : String;
                AXE                : String;
                AXEPLUS            : Array[1..5] of String[2] ;
                BUDSECT            : String;
                DEBIT              : Double;
                CREDIT             : Double;
                CONFIDENTIEL       : String ;
                REFINTERNE         : String ;
                ETABLISSEMENT      : String3;
                QUALIFPIECE        : String ;
                DEBITEURO          : Double;
                CREDITEURO         : Double;
                CHRONO             : Integer ;
                CHAMP              : String ;
                MODIF              : Boolean ;
                INDICETABLE        : Byte;
                End ;
type TFDevise = Class
  Code : String3 ;
  Decimale : Integer ;
  Quotite : Double ;
  TauxDev : Double ;
  END ;

type String5 = String[5] ;


type TAnal = RECORD
   Journal,NatPiece,QualP : String3 ;
   General : String17 ;
   NumP,Numl,Numv : Integer ;
   RefInt,Lib : String ;
   Debit,Credit  : Double ;
   END ;

const SepLigneIE=Chr(165) ;
Const Sn2Orli : Boolean = FALSE ;

Function RecupDevise(Code : String3 ; var DecDev : Integer ; var Quotite : Double ; TDev : TList) : Boolean ;

implementation

Function RecupDevise(Code : String3 ; var DecDev : Integer ; var Quotite : Double ; TDev : TList) : Boolean ;
var TDevise : TFDevise ;
    i : integer ;
BEGIN
Result:=FALSE ;
if (Code='') or (Code=V_PGI.DevisePivot) then
  BEGIN
  DecDev:=V_PGI.OkDecV ; Quotite:=1 ;
  END else
  BEGIN
  for i:=0 to TDev.Count-1 do
    BEGIN
    TDevise:=TDev[i] ;
    if (TDevise.Code=Code) then
      BEGIN
      DecDev:=TDevise.Decimale ; Quotite:=TDevise.Quotite ;
      Result:=TRUE ;
      Break ;
      END ;
    END ;
  END ;
END ;

end.
