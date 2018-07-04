unit CPProcMetier;

interface

uses
  uTob, HEnt1, CPTypeCons, HCtrls, StdCtrls;


function  GetTenuEuro : Boolean;

{Conversion Nature, DataType, FicheirBase}
function  GetSousPlanAxe  (fb : TFichierBase ; vNumPlan : integer) : TSousPlan ;
function  FbSousPlantoFb  (fb : TFichierBase) : TFichierBase ;
function  FbToAxe         (fb : TFichierBase) : string ;
function  tzToNature      (tz : TZoomTable  ) : string ;
function  tzToChampNature (tz : TZoomTable; AvecPrefixe : Boolean) : string;
function  NatureToTz      (Nat : string) : TZoomTable;
function  AxeToFb         (Axe : string) : TFichierBase;
function  AxeToFbBud      (Axe : string) : TFichierBase;
function  AxeToTz         (Axe : string) : TZoomTable ;
function  AxeToDataType   (CodeAxe : string) : string;
function  NatureToDataType(CodeTable : string) : string;

{Caractères de bourrage}
function  BourreLaDoncSurLaTable   (Code, St : string; ForceLg : Integer = 0 ) : string;
function  BourreLaDoncSurLesComptes(st : string; c : string = '') : string;
function  BourreLaDonc             (St : string; LeType : TFichierBase ) : string;
function  BourreEtLess             (St : string; LeType : TFichierBase; ForceLg : Integer = 0) : string;
function  BourreLess               (St : string; LeType : TFichierBase) : string;
function  BourreOuTronque          (St : string; fb : TFichierBase) : string;
function  GetBourreLibre           : Char;

{TVA et TPF}
function  TVA2ENCAIS(ModeTVA, Tva : String3; Achat : Boolean; RG : Boolean = False; FarFae : Boolean = False) : string ;
function  TPF2ENCAIS(ModeTVA, Tpf : String3; Achat : Boolean; FarFae : Boolean = False) : string;
function  TVA2CPTE  (ModeTVA, Tva : String3; Achat : Boolean; FarFae : Boolean = False) : string ;
function  TPF2CPTE  (ModeTVA, Tpf : String3; Achat : Boolean; FarFae : Boolean = False) : string ;
function  TVA2TAUX  (ModeTVA, Tva : String3; Achat : Boolean) : Real;
function  TPF2TAUX  (ModeTVA, Tpf : String3; Achat : Boolean) : Real;
function  HT2TPF    (THT : Real; ModeTVA : String3; SoumisTPF :  Boolean; Tva, Tpf : String3; Achat : Boolean; Dec : Integer) : Real;
function  HT2TVA    (THT : Real; ModeTVA : String3; SoumisTPF :  Boolean; Tva, Tpf : String3; Achat : Boolean; Dec : Integer) : Real;
function  TTC2HT    (TTTC : Real; ModeTVA : String3; SoumisTPF :  Boolean; Tva, Tpf : String3; Achat : Boolean; Dec : Integer) : Real;
function  CPGetMontantTVA(vStCodeTva : string ; vDateDebut, vDateFin : TDateTime) : Double;

{Tables Libres}
procedure Dechiffre(St : string; var C : TFF);
procedure GetLibelleTableLibre(Pref : string; LesLib : HTStringList);
function  CompareTL(St1, St2 : string) : Boolean;
procedure ChargeComboTableLibre(Cod : string; DesValues, DesItems : HTStrings);

{Divers....}
procedure CorrespToCodes(Plan : THValComboBox; var C1, C2 : TComboBox);
procedure AfficheLeSolde (T : THNumEdit; TD, TC : Double);

function  EstConfidentiel (Conf : string) : Boolean;
function  EstSQLConfidentiel(StTable : string; Cpte : String17) : Boolean;

function  _ChampExiste(NomTable, NomChamp : string) : Boolean;
function  SQLPremierDernier(fb : TFichierbase; Prem : Boolean) : string;
procedure PremierDernierRub(ZoomTable : TZoomTable; SynPlus : string; var Cpt1, Cpt2 : string);
procedure PremierDernier(fb : TFichierBase; var Cpt1, Cpt2 : string);

function  PieceSurFolio(Jal : string) : Boolean;
function  AuMoinsUneImmo : Boolean;
procedure CreerDeviseTenue(LaMonnaie : string);
Procedure CPInitSequenceCompta;


implementation

uses
  {$IFNDEF EAGLCLIENT}
  MajTable,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF EAGLCLIENT}
  UtilPgi, Ent1, SysUtils, ParamSoc,
  {$IFNDEF EAGLSERVER}
  {$IFDEF COMPTA}
  galOutil,
  {$ENDIF}
  {$ENDIF EAGLSERVER}
  {$IFDEF NOVH}
  ULibCpContexte,
  {$ENDIF NOVH}

  CPVersion, CPProcGen;


{---------------------------------------------------------------------------------------}
function GetTenuEuro : Boolean;
{---------------------------------------------------------------------------------------}
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','GetTenueEuro') ;
{$ENDIF}
{$IFDEF NOVH}
 Result := GetParamsoc ('SO_TENUEEURO');
{$ELSE}
 Result := VH^.TenueEuro;
{$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function BourreLaDonc ( St : String ; LeType : TFichierBase ) : string ;
{---------------------------------------------------------------------------------------}
var Lg,ll,i : Integer ;
    Bourre  : Char ;
    lInfoCpta : TInfoCpta ;
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','BourreLaDonc') ;
  {$ENDIF}
  If LeType In [fbAxe1..fbAux,fbNatCpt] Then
   BEGIN
   lInfoCpta := GetInfoCpta(leType) ;
   Lg := lInfoCpta.lg ;
   If LeType=fbNatCpt Then Bourre:=GetBourreLibre Else Bourre:=lInfoCpta.Cb ;
   Result:=St ; ll:=Length(Result) ;
   If ll<Lg then for i:=ll+1 to Lg do Result:=Result+Bourre ;
   END Else Result:=St ;
end ;

{---------------------------------------------------------------------------------------}
function BourreEtLess (St : string; LeType : TFichierBase; ForceLg : Integer = 0) : string;
{---------------------------------------------------------------------------------------}
var Lg,ll,i : Integer ;
    Bourre  : Char ;
     lInfoCpta : TInfoCpta ;
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','BourreEtLess') ;
  {$ENDIF}
  If LeType In [fbAxe1..fbAux,fbNatCpt] Then
   BEGIN
   lInfoCpta := GetInfoCpta(leType) ;
   Lg := lInfoCpta.lg ; If ForceLg>0 Then Lg:=ForceLg ;
   If LeType=fbNatCpt Then Bourre:=GetBourreLibre Else Bourre:=lInfoCpta.Cb ;
   If Length(St)>Lg Then St:=Trim(Copy(St,1,Lg)) ;
   Result:=St ; ll:=Length(Result) ;
   If ll<Lg then
      BEGIN
      for i:=ll+1 to Lg do Result:=Result+Bourre ;
      END Else Result:=Copy(Result,1,Lg) ;
   END Else Result:=St ;
end;

{---------------------------------------------------------------------------------------}
function BourreLess(St : string; LeType : TFichierBase) : string;
{---------------------------------------------------------------------------------------}
var Lg,ll,i : Integer ;
    Bourre  : Char ;
    lInfoCpta : TInfoCpta ;
    Label 0 ;
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','BourreLess') ;
  {$ENDIF}
  lInfoCpta := GetInfoCpta(leType) ;
  Lg := lInfoCpta.lg ; Bourre:=lInfoCpta.Cb ;
  ll:=Length(St) ; If ll>Lg Then St:=Copy(St,1,Lg) ; i:=Length(St) ;
  While i>0 Do BEGIN If St[i]=Bourre Then St[i]:=' ' Else Goto 0 ; Dec(i) ; END ;
  0:Result:=Trim(St) ;
end ;

{---------------------------------------------------------------------------------------}
function BourreLaDoncSurLesComptes(st : string ; c : string = '') : string;
{---------------------------------------------------------------------------------------}
var st1 : string ;
    l : integer;
    cb : string;
    lInfoCpta : TInfoCpta ;
begin
  lInfoCpta := GetInfoCpta(fbGene) ;
  if c = '' then cb := lInfoCpta.Cb else cb := c;
  st1 := Trim(st) ;
  l := lInfoCpta.Lg;
  if l<= 0 then {!!!!????} else
  if l<Length(st1) then st1 := Copy(st1,1,l);
  while Length(st1)<l do st1 := st1 + cb ;
  Result := st1 ;
end;

{---------------------------------------------------------------------------------------}
function BourreOuTronque(St : string; fb : TFichierBase) : string;
{---------------------------------------------------------------------------------------}

    {----------------------------------------------------------------------}
    function LongMax2 ( fb : TFichierBase ) : integer ;
    {----------------------------------------------------------------------}
    Var
      i : Integer ;
    BEGIN
    if fb in [fbAxe1..fbAux] then Result:=GetInfoCpta(fb).lg else
    Case fb of
      fbJal     : Result:=3 ;
      fbAxe1SP1..fbAxe5SP6 : BEGIN      { DONE -olaurent -cdll serveur : a reecrire }
                             i:=((Byte(fb)-Byte(fbAxe1SP1)) Mod 6)+1 ;
                             Result:=GetSousPlanAxe(fbSousPlanTofb(fb),i).Longueur ;
                             END ;
      else Result:=17 ;
      end ;
    END ;

var
  Diff : Integer ;
BEGIN
  {$IFDEF TTW}
   cWA.MessagesAuClient('COMSX.IMPORT','','BourreOuTronque') ;
  {$ENDIF}
  St:=Trim(St) ; Result:=St ; if (St='') then Exit ;
  Diff:=Length(St)-LongMax2(fb) ;
  if Diff>0 then St:=BourreLess(St,fb) ;
  Result:=BourreLaDonc(St,fb) ;
END ;

{---------------------------------------------------------------------------------------}
Function BourreLaDoncSurLaTable ( Code, St : String; ForceLg : Integer = 0) : string ;
{---------------------------------------------------------------------------------------}
var Lg,ll,i : Integer ;
{$IFNDEF EAGLSERVER}
 j : integer ;
{$ENDIF}
    Bourre    : Char ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','BourreLaDoncSurLaTable ') ;
{$ENDIF}
Result:=St ;
Case Code[1] Of
     'G' : i:=1 ; 'T' : i:=2 ; 'S' : i:=3 ; 'B' : i:=4 ;
     'D' : i:=5 ; 'E' : i:=6 ; 'A' : i:=7 ; 'U' : i:=8 ;
     'I' : i:=9 ;
     Else i:=0 ;
   END ;
{$IFDEF EAGLSERVER}
If i>0 Then Lg:=ForceLg Else exit ;
{$ELSE}
{$IFNDEF NOVH}
{ DONE -olaurent -cdll serveur : a faire }
If i>0 Then BEGIN j:=StrToInt(Copy(Code,2,2))+1 ; Lg:=GetLgTableLibre(i,j) ; END Else exit ;
{$ENDIF}
{$ENDIF}
Bourre:=GetBourreLibre ;
ll:=Length(Result) ;
If ll<Lg then for i:=ll+1 to Lg do Result:=Result+Bourre ;
end ;

{---------------------------------------------------------------------------------------}
Function GetBourreLibre : Char;
{---------------------------------------------------------------------------------------}
{$IFDEF NOVH}
var
   BLibre : String;
{$ENDIF}
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','GetBourreLibre') ;
  {$ENDIF}

  {$IFDEF NOVH}
  Result := #0 ;
  BLibre := GetParamSocSecur('SO_BOURRELIB', TRUE);
  if BLibre<>'' then Result := BLibre[1] ;
  {$ELSE}
  Result := VH^.BourreLibre;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function GetSousPlanAxe(fb : TFichierBase ; vNumPlan : integer) : TSousPlan ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','GetSousPlanAxe') ;
  {$ENDIF}
  {$IFDEF NOVH}
  result := TCPContexte.GetCurrent.InfoCpta.SousPlanAxe[fb,vNumPlan] ;
  {$ELSE}
  result := VH^.SousPlanAxe[fb,vNumPlan] ;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
Function FbSousPlantoFb ( fb : TFichierBase ) : TFichierBase ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','FbSousPlantoFb');
  {$ENDIF}
  Result:=fbAxe1 ;
  case fb of
    fbAxe1SP1..fbAxe1SP6 : Result := fbAxe1;
    fbAxe2SP1..fbAxe2SP6 : Result := fbAxe2;
    fbAxe3SP1..fbAxe3SP6 : Result := fbAxe3;
    fbAxe4SP1..fbAxe4SP6 : Result := fbAxe4;
    fbAxe5SP1..fbAxe5SP6 : Result := fbAxe5;
  end;
end;

{---------------------------------------------------------------------------------------}
function  FbToAxe ( fb : TFichierBase ) : string;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TTW}
   cWA.MessagesAuClient('COMSX.IMPORT','','FbToAxe') ;
  {$ENDIF}
       if fb in [fbBudSec1] then Result := 'A1'
  else if fb in [fbBudSec2] then Result := 'A2'
  else if fb in [fbBudSec3] then Result := 'A3'
  else if fb in [fbBudSec4] then Result := 'A4'
  else if fb in [fbBudSec5] then Result := 'A5'
                            else Result := 'A' + Chr(49 + Byte(fb));
end;

{---------------------------------------------------------------------------------------}
function  AxeToFb (Axe : string) : TFichierBase;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','AxeToFb') ;
  {$ENDIF}
  Result := fbAxe1;
  if Length(Axe) >= 2 then Result := TFichierBase(Ord(Axe[2]) - 49);
end;

{---------------------------------------------------------------------------------------}
function  AxeToFbBud(Axe : string) : TFichierBase;
{---------------------------------------------------------------------------------------}
begin
  Result := fbBudsec1 ;
  if Length(Axe) >= 2 then
    case Axe[2] of
      '1' : Result := fbBudsec1;
      '2' : Result := fbBudsec2;
      '3' : Result := fbBudsec3;
      '4' : Result := fbBudsec4;
      '5' : Result := fbBudsec5;
    end;
end;

{---------------------------------------------------------------------------------------}
function  tzToNature(tz : TZoomTable) : string;
{---------------------------------------------------------------------------------------}
begin
  case tz of
    tzNatGene0,tzNatGene1,tzNatGene2,tzNatGene3,tzNatGene4,tzNatGene5,
    tzNatGene6,tzNatGene7,tzNatGene8,tzNatGene9 : Result:='G'+FormatFloat('00',Byte(tz)-Byte(tzNatGene0)) ;
    tzNatTiers0,tzNatTiers1,tzNatTiers2,tzNatTiers3,tzNatTiers4,tzNatTiers5,
    tzNatTiers6,tzNatTiers7,tzNatTiers8,tzNatTiers9 : Result:='T'+FormatFloat('00',Byte(tz)-Byte(tzNatTiers0)) ;
    tzNatSect0,tzNatSect1,tzNatSect2,tzNatSect3,tzNatSect4,tzNatSect5,
    tzNatSect6,tzNatSect7,tzNatSect8,tzNatSect9 : Result:='S'+FormatFloat('00',Byte(tz)-Byte(tzNatSect0)) ;
    tzNatBud0,tzNatBud1,tzNatBud2,tzNatBud3,tzNatBud4,tzNatBud5,
    tzNatBud6,tzNatBud7,tzNatBud8,tzNatBud9 : Result:='B'+FormatFloat('00',Byte(tz)-Byte(tzNatBud0)) ;
    tzNatBudS0,tzNatBudS1,tzNatBudS2,tzNatBudS3,tzNatBudS4,tzNatBudS5,
    tzNatBudS6,tzNatBudS7,tzNatBudS8,tzNatBudS9 : Result:='D'+FormatFloat('00',Byte(tz)-Byte(tzNatBudS0)) ;
    tzNatEcrE0,tzNatEcrE1,tzNatEcrE2,tzNatEcrE3 : Result:='E'+FormatFloat('00',Byte(tz)-Byte(tzNatEcrE0)) ;
    tzNatEcrA0,tzNatEcrA1,tzNatEcrA2,tzNatEcrA3 : Result:='A'+FormatFloat('00',Byte(tz)-Byte(tzNatEcrA0)) ;
    tzNatEcrU0,tzNatEcrU1,tzNatEcrU2,tzNatEcrU3 : Result:='U'+FormatFloat('00',Byte(tz)-Byte(tzNatEcrU0)) ;
    tzNatImmo0,tzNatImmo1,tzNatImmo2,tzNatImmo3,tzNatImmo4,tzNatImmo5,
    tzNatImmo6,tzNatImmo7,tzNatImmo8,tzNatImmo9 : Result:='I'+FormatFloat('00',Byte(tz)-Byte(tzNatImmo0)) ;
  end;
end;

{JP 13/06/05 : Equivaut à NatureToTz pour les DataTypes
{---------------------------------------------------------------------------------------}
function NatureToDataType(CodeTable : string) : string;
{---------------------------------------------------------------------------------------}
begin
  case CodeTable[1] of
    'B' : Result:= 'TZNATBUD'   + CodeTable[3];
    'G' : Result:= 'TZNATGENE'  + CodeTable[3];
    'S' : Result:= 'TZNATSECT'  + CodeTable[3];
    'T' : Result:= 'TZNATTIERS' + CodeTable[3];
    'D' : Result:= 'TZNATBUDS'  + CodeTable[3];
    'E' : Result:= 'TZNATECR'   + CodeTable[3];
    'A' : Result:= 'TZNATECRA'  + CodeTable[3];
    'U' : Result:= 'TZNATECRU'  + CodeTable[3];
    'I' : Result:= 'TINATIMMO'  + CodeTable[3];
  end;
end;

{JP 13/06/05 : Equivaut à AxeToTz pour les DataTypes
{---------------------------------------------------------------------------------------}
function AxeToDataType(CodeAxe : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := 'TZSECTION' + CodeAxe[2];
end;

{---------------------------------------------------------------------------------------}
function  AxeToTz(Axe : string) : TZoomTable;
{---------------------------------------------------------------------------------------}
begin
  Result := TZoomTable(Ord(tzSection) + Ord(Axe[2]) - 49);
end;

{---------------------------------------------------------------------------------------}
function tzToChampNature( tz : TZoomTable; AvecPrefixe : Boolean) : string;
{---------------------------------------------------------------------------------------}
var
  StPref,
  StNum : string;
begin
  StPref := '';
  STNum  := '0';
  
  case tz of
    tzNatGene0,tzNatGene1,tzNatGene2,tzNatGene3,tzNatGene4,tzNatGene5,
    tzNatGene6,tzNatGene7,tzNatGene8,tzNatGene9 : begin StPref:='G_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatGene0)) ; end ;
    tzNatTiers0,tzNatTiers1,tzNatTiers2,tzNatTiers3,tzNatTiers4,tzNatTiers5,
    tzNatTiers6,tzNatTiers7,tzNatTiers8,tzNatTiers9 : begin StPref:='T_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatTiers0)) ; end ;
    tzNatSect0,tzNatSect1,tzNatSect2,tzNatSect3,tzNatSect4,tzNatSect5,
    tzNatSect6,tzNatSect7,tzNatSect8,tzNatSect9 : begin StPref:='S_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatSect0)) ; end ;
    tzNatBud0,tzNatBud1,tzNatBud2,tzNatBud3,tzNatBud4,tzNatBud5,
    tzNatBud6,tzNatBud7,tzNatBud8,tzNatBud9 : begin StPref:='B_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatBud0)) ; end ;
    tzNatBudS0,tzNatBudS1,tzNatBudS2,tzNatBudS3,tzNatBudS4,tzNatBudS5,
    tzNatBudS6,tzNatBudS7,tzNatBudS8,tzNatBudS9 : begin StPref:='D_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatBudS0)) ; end ;
    tzNatEcrE0,tzNatEcrE1,tzNatEcrE2,tzNatEcrE3 : begin StPref:='E_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatEcrE0)) ; end ;
    tzNatEcrA0,tzNatEcrA1,tzNatEcrA2,tzNatEcrA3 : begin StPref:='Y_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatEcrA0)) ; end ;
    tzNatEcrU0,tzNatEcrU1,tzNatEcrU2,tzNatEcrU3 : begin StPref:='U_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatEcrU0)) ; end ;
    tzNatImmo0,tzNatImmo1,tzNatImmo2,tzNatImmo3,tzNatImmo4,tzNatImmo5,
    tzNatImmo6,tzNatImmo7,tzNatImmo8,tzNatImmo9 : begin StPref:='I_' ; StNum:=FormatFloat('0',Byte(tz)-Byte(tzNatImmo0)) ; end ;
    tzRien : begin result:='' ; Exit ; end ;
  end;

  if not AvecPrefixe then StPref := '';

  Result := StPref + 'TABLE' + StNum;
end;

{---------------------------------------------------------------------------------------}
function NatureToTz (Nat : string) : TZoomTable;
{---------------------------------------------------------------------------------------}
begin
  case Nat[1] Of
    'B' : Result := TZoomtable(Ord(tzNatBud0  ) +Ord(Nat[3]) - 48);
    'G' : Result := TZoomtable(Ord(tzNatGene0 ) +Ord(Nat[3]) - 48);
    'S' : Result := TZoomtable(Ord(tzNatSect0 ) +Ord(Nat[3]) - 48);
    'T' : Result := TZoomtable(Ord(tzNatTiers0) +Ord(Nat[3]) - 48);
    'D' : Result := TZoomtable(Ord(tzNatBudS0 ) +Ord(Nat[3]) - 48);
    'E' : Result := TZoomtable(Ord(tzNatEcrE0 ) +Ord(Nat[3]) - 48);
    'A' : Result := TZoomtable(Ord(tzNatEcrA0 ) +Ord(Nat[3]) - 48);
    'U' : Result := TZoomtable(Ord(tzNatEcrU0 ) +Ord(Nat[3]) - 48);
    'I' : Result := TZoomtable(Ord(tzNatImmo0 ) +Ord(Nat[3]) - 48);
  else
    Result := tzRien;
  end;
end;

{---------------------------------------------------------------------------------------}
function TVA2ENCAIS(ModeTVA,Tva : String3 ; Achat : Boolean; RG : boolean=false; FarFae : boolean=false) : String ;
{---------------------------------------------------------------------------------------}
Var TOBT : TOB ;
BEGIN
  Result:='' ;
  {$IFNDEF NOVH}
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTVA,ModeTVA,TVA],False) ;
  {$ENDIF NOVH}
  if TOBT<>Nil then
  BEGIN
    if Achat then
    begin
      if not RG then
      begin
        if not FarFae then
          Result := TOBT.GetValue('TV_ENCAISACH')
          else
          Result := TOBT.GetValue('TV_CPTACHFARFAE')
      end else
        Result:=TOBT.GetValue('TV_CPTACHRG');
    end else
    begin
      if not RG then
      begin
        if not FarFae then
          Result:=TOBT.GetValue('TV_ENCAISVTE')
          else
          Result := TOBT.GetValue('TV_CPTVTEFARFAE')
      end else
        Result:=TOBT.GetValue('TV_CPTVTERG')
    end;
  END ;
END ;

{---------------------------------------------------------------------------------------}
function HT2TPF ( THT : Real ; ModeTVA : string3 ; SoumisTPF :  boolean ; Tva,Tpf : String3 ; Achat : Boolean ; Dec : integer) : Real ;
{---------------------------------------------------------------------------------------}
begin
  HT2Tpf:=0 ; if Not SoumisTpf then Exit ;
  if QuelPaysLocalisation=CodeISOES then //XMG 24/02/2005
     HT2TPF:=Arrondi(THT*(Tpf2Taux(ModeTVA,Tpf,Achat)),Dec)
  {$IFNDEF NOVH}
  else
  if VH^.TpfDehors then
     BEGIN
     if VH^.TpfApres then HT2TPF:=Arrondi(THT*(1.0+Tva2Taux(ModeTVA,Tva,Achat))*Tpf2Taux(ModeTVA,Tpf,Achat),Dec)
                     else HT2TPF:=Arrondi(THT*Tpf2Taux(ModeTVA,Tpf,Achat),Dec) ;
     END else
     BEGIN
     HT2TPF:=Arrondi(THT*(1.0+Tpf2Taux(ModeTVA,Tpf,Achat))*Tpf2Taux(ModeTVA,Tpf,Achat),Dec) ;
     END ;
  {$ENDIF NOVH}
END ;

{---------------------------------------------------------------------------------------}
function HT2TVA (THT : Real; ModeTVA : String3; SoumisTPF :  Boolean; Tva, Tpf : String3; Achat : Boolean; Dec : Integer) : Real;
{---------------------------------------------------------------------------------------}
BEGIN
{$IFDEF NOVH}
HT2TVA := 0;
{$ENDIF}
if (QuelPaysLocalisation=CodeISOES) or (Not SoumisTPF) then //XMG 24/02/2005
   BEGIN
   HT2TVA:=Arrondi(THT*Tva2Taux(ModeTVA,Tva,Achat),Dec) ;
   END else
   BEGIN
  {$IFNDEF NOVH}
   if VH^.TpfDehors then
      BEGIN
      if VH^.TpfApres then HT2TVA:=Arrondi(THT*Tva2Taux(ModeTVA,Tva,Achat),Dec)
                     else HT2TVA:=Arrondi(THT*((1.0+Tpf2Taux(ModeTVA,Tpf,Achat)))*Tva2Taux(ModeTVA,Tva,Achat),Dec) ;
      END else
      BEGIN
      HT2TVA:=Arrondi(THT*(((1.0+Tpf2Taux(ModeTVA,Tpf,Achat))*Tpf2Taux(ModeTVA,Tpf,Achat))+1)*Tva2Taux(ModeTVA,Tva,Achat),Dec) ;
      END ;
  {$ENDIF NOVH}
   END ;
END ;

{---------------------------------------------------------------------------------------}
function TTC2HT(TTTC : Real; ModeTVA : String3; SoumisTPF :  Boolean; Tva, Tpf : String3; Achat : Boolean; Dec : Integer) : Real;
{---------------------------------------------------------------------------------------}
Var
  IntFR,CoefTpf : real ;
BEGIN
{$IFDEF NOVH}
TTC2HT := 0;
{$ENDIF}
if Not SoumisTPF then
   BEGIN
   TTC2HT:=Arrondi(TTTC/(1.0+Tva2Taux(ModeTva,Tva,Achat)),Dec) ;
   END else
   BEGIN
   if QuelPaysLocalisation=CodeISOES then //XMG 24/02/2005
      TTC2HT:=Arrondi(TTTC/(1.0+(Tva2Taux(ModeTva,Tva,Achat)+Tpf2Taux(ModeTVA,Tpf,Achat))),Dec)
  {$IFNDEF NOVH}
   else
   if VH^.TpfDehors then
      BEGIN
      if VH^.TpfApres then
         BEGIN
         TTC2HT:=Arrondi((TTTC)/(1.0+Tva2Taux(ModeTVA,Tva,Achat))/(1.0+Tpf2Taux(ModeTVA,Tpf,Achat)),Dec) ;
         END else
         BEGIN
         INTFR:=(1.0+Tva2Taux(ModeTVA,Tva,Achat)) ;
         TTC2HT:=Arrondi((TTTC-IntFR)/(1.0+Tva2Taux(ModeTVA,Tva,Achat))/(1.0+Tpf2Taux(ModeTVA,Tpf,Achat)),Dec) ;
         END ;
      END else
      BEGIN
      CoefTPF:=1.0+(Tpf2Taux(ModeTVA,Tpf,Achat)*(1.0+Tpf2Taux(ModeTVA,Tpf,Achat))) ;
      TTC2HT:=Arrondi(TTTC/(1.0+Tva2Taux(ModeTVA,Tva,Achat))/CoefTPF,Dec) ;
      END ;
{$ENDIF}
   END ;
END ;

{---------------------------------------------------------------------------------------}
function CPGetMontantTVA(vStCodeTva : string ; vDateDebut, vDateFin : TDateTime) : Double;
{---------------------------------------------------------------------------------------}
var lQuery    : TQuery;
    lTauxTva  : Double;
    lStWhere  : string;
    lStAxeTVa : string;
    lBoTvaSurEnc : Boolean;
begin
  Result := 0;
  if not CPEstTvaActivee then Exit;

  lStAxeTva  := GetParamSocSecur('SO_CPPCLAXETVA', '');
  if lStAxeTva = '' then
  begin
    lStAxeTva := GetColonneSQL('AXE', 'X_AXE', 'X_LIBELLE = "TVA"');
    SetParamSoc('SO_CPPCLAXETVA', lStAxeTva);
  end;

  // Contrôle OK, on calcule

  lBoTvaSurEnc := GetParamSocSecur('SO_OUITVAENC', False);

  {$IFDEF EAGLSERVER}
  lTauxTva := 0;
  {$ELSE}
  // Recherche du Taux dans CHOIXDPSTD
  lTauxTva := Valeur(GetColonneSql('CHOIXDPSTD', 'YDS_LIBRE', 'YDS_TYPE = "TVA" AND ' +
                            'YDS_CODE = "' + vStCodeTva + '" AND ' +
                            'YDS_NODOSSIER = "' + NoDossierBaseCommune + '" AND ' +
                            'YDS_PREDEFINI = "CEG"'));
  {$ENDIF EAGLSERVER}

  lStWhere := 'Y_SECTION = "' +  vStCodeTva + '" AND ' +
              'Y_AXE = "' + lStAxeTva + '" AND ' +
              'Y_DATECOMPTABLE >= "' + UsDateTime(vDateDebut) + '" AND ' +
              'Y_DATECOMPTABLE <= "' + UsDateTime(VDateFin) + '"' ;

  try
    lQuery := OpenSQL('SELECT SUM(Y_CREDIT)-SUM(Y_DEBIT) TOTAL FROM ANALYTIQ WHERE ' + lStWhere, True);

    Result := lQuery.FindField('TOTAL').AsFloat;

    if lBoTvaSurEnc and (not GetParamSocSecur('SO_CPTVARECDEP',False)) then
      Result := Result / (1 + (lTauxTva /100));

  finally
    Ferme(lQuery);
  end;
end;

{---------------------------------------------------------------------------------------}
FUNCTION TVA2CPTE(ModeTVA,Tva : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
{---------------------------------------------------------------------------------------}
Var TOBT : TOB ;
BEGIN
  Result:='' ;
  {$IFNDEF NOVH}
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTVA,ModeTVA,TVA],False) ;
  {$ENDIF NOVH}
  if TOBT<>Nil then
  BEGIN
    if Achat then
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_CPTEACH')
        else
        Result:=TOBT.GetValue('TV_CPTACHFARFAE');
    end else
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_CPTEVTE')
        else
        Result:=TOBT.GetValue('TV_CPTVTEFARFAE');
    end;
  END ;
END ;

{---------------------------------------------------------------------------------------}
function TVA2TAUX(ModeTVA,Tva : String3 ; Achat : Boolean) : Real ;
{---------------------------------------------------------------------------------------}
Var TOBT : TOB ;
BEGIN
Result:=0 ;
{$IFNDEF NOVH}
TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTVA,ModeTVA,TVA],False) ;
{$ENDIF}
if TOBT<>Nil then
    BEGIN
    if Achat then Result:=TOBT.GetValue('TV_TAUXACH')/100.0
             else Result:=TOBT.GetValue('TV_TAUXVTE')/100.0 ;
    END ;
END ;

{---------------------------------------------------------------------------------------}
function  TPF2CPTE(ModeTVA,Tpf : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
{---------------------------------------------------------------------------------------}
Var TOBT : TOB ;
BEGIN
  Result:='' ;
  {$IFNDEF NOVH}
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTPF,ModeTVA,TPF],False) ;
  {$ENDIF}
  if TOBT<>Nil then
  BEGIN
    if Achat then
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_CPTEACH')
        else
        Result := TOBT.GetValue('TV_CPTACHFARFAE');
    end else
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_CPTEVTE')
        else
        Result := TOBT.GetValue('TV_CPTVTEFARFAE');
    end;
  END ;
END ;

{---------------------------------------------------------------------------------------}
FUNCTION TPF2TAUX(ModeTVA,Tpf : String3 ; Achat : Boolean) : Real ;
{---------------------------------------------------------------------------------------}
Var TOBT : TOB ;
BEGIN
Result:=0 ;
{$IFNDEF NOVH}
TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTPF,ModeTVA,TPF],False) ;
{$ENDIF}
if TOBT<>Nil then
    BEGIN
    if Achat then Result:=TOBT.GetValue('TV_TAUXACH')/100.0
             else Result:=TOBT.GetValue('TV_TAUXVTE')/100.0 ;
    END ;
END ;

{---------------------------------------------------------------------------------------}
function TPF2ENCAIS(ModeTVA,Tpf : String3 ; Achat : Boolean; FarFae : boolean=false) : String ;
{---------------------------------------------------------------------------------------}
Var TOBT : TOB ;
BEGIN
  Result:='' ;
  {$IFNDEF NOVH}
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[VH^.DefCatTPF,ModeTVA,TPF],False) ;
  {$ENDIF}
  if TOBT<>Nil then
  BEGIN
    if Achat then
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_ENCAISACH')
        else
        Result := TOBT.GetValue('TV_CPTACHFARFAE');
    end else
    begin
      if not FarFae then
        Result:=TOBT.GetValue('TV_ENCAISVTE')
        else
        Result := TOBT.GetValue('TV_CPTVTEFARFAE');
    end;
  END ;
END ;

{---------------------------------------------------------------------------------------}
Procedure GetLibelleTableLibre(Pref : String ; LesLib : HTStringList) ;
{---------------------------------------------------------------------------------------}
Var QLoc : TQuery ;
    lStDossier : string ;
BEGIN
// Gestion du multi-dossier
  if EstTablePartagee( 'NATCPTE' )
    then lStDossier := TableToBase( 'NATCPTE' )
    else lStDossier := '' ;
QLoc:=OpenSelect('Select CC_LIBELLE,CC_ABREGE,CC_CODE From CHOIXCOD Where CC_TYPE="NAT" And CC_CODE Like"'+Pref+'%" Order by CC_CODE',lStDossier) ;
LesLib.Clear ;
While Not QLoc.Eof do
   BEGIN LesLib.Add(QLoc.Fields[0].AsString+';'+QLoc.Fields[1].AsString) ; QLoc.Next ; END ;
Ferme(QLoc) ;
END ;

{---------------------------------------------------------------------------------------}
Function CompareTL(St1,St2 : String) : Boolean ;
{---------------------------------------------------------------------------------------}
Var C1,C2 : TFF ;
    i : Integer ;
    Tous1,Tous2,OkOk : Boolean ;
BEGIN
Result:=TRUE ;
Fillchar(C1,SizeOf(C1),#0) ; Fillchar(C2,SizeOf(C2),#0) ;
Dechiffre(St1,C1) ; Dechiffre(St2,C2) ;
For i:=1 To 10 Do
  BEGIN
  If C1[i].Deb='' Then Continue ; If C2[i].Deb='' Then Continue ;
  Tous1:=C1[i].Deb='*' ; Tous2:=C2[i].Deb='*' ;
  If Tous1 And (C2[i].Deb<>'') Then Exit Else If Tous2 And (C1[i].Deb<>'') Then Exit Else
    BEGIN
    If (Pos('###',C1[i].Deb)>0) Or (Pos('---',C1[i].Deb)>0) Or
       (Pos('###',C2[i].Deb)>0) Or (Pos('---',C2[i].Deb)>0) Then Continue ;
    OkOk:=((C1[i].Deb<C2[i].Deb) And (C1[i].Fin<C2[i].Deb)) Or
          ((C2[i].Deb<C1[i].Deb) And (C2[i].Fin<C1[i].Deb)) ;
    If Not OkOk THen Exit ;
    END ;
  END ;
Result:=FALSE ;
END ;

{---------------------------------------------------------------------------------------}
procedure ChargeComboTableLibre(Cod : string; DesValues, DesItems : HTStrings);
{---------------------------------------------------------------------------------------}
Var QLoc : TQuery ;
    lStDossier : string ;
BEGIN
DesValues.Clear ; DesItems.Clear ;
if Cod='' then Exit ;
// Gestion du multi-dossier
  if EstTablePartagee( 'NATCPTE' )
    then lStDossier := TableToBase( 'NATCPTE' )
    else lStDossier := '' ;
QLoc:=OpenSelect('Select CC_CODE,CC_LIBELLE From CHOIXCOD Where CC_TYPE="NAT" And CC_CODE Like "'+Cod+'%"',lStDossier, True, 'CHOIXCOD') ;
While Not QLoc.Eof do
   BEGIN
   DesValues.Add(QLoc.Fields[0].AsString) ; DesItems.Add(QLoc.Fields[1].AsString) ;
   QLoc.Next ;
   END ;
Ferme(QLoc) ;
END ;

{---------------------------------------------------------------------------------------}
procedure Dechiffre(St : string ; var C : TFF) ;
{---------------------------------------------------------------------------------------}
Var St1,St2 : String ;
    SDeb,SFin : String ;
    i : Integer ;
    ATraiter : Boolean ;
    Tous : Boolean ;
BEGIN
i:=Pos(':',St) ;
If i>0 Then BEGIN St1:=Copy(St,1,i-1) ; St2:=Copy(St,i+1,Length(St)-i) END Else
  BEGIN St1:=St ; St2:=St ; END ;
For i:=1 To 10 Do
  BEGIN
  SDeb:=ReadTokenstV(St1) ; SFin:=ReadTokenstV(St2) ;
  ATRaiter:=(Pos('-',SDeb)<=0) And (Pos('#',SDeb)<=0) ; Tous:=Pos('*',SDeb)>0 ;
  If ATraiter Then
    BEGIN
    If Tous Then BEGIN C[i].Deb:='*' ; C[i].Fin:='*' ; End
            Else BEGIN C[i].Deb:=SDeb ; C[i].Fin:=SFin ; END ;
    END ;
  END ;
END ;

{---------------------------------------------------------------------------------------}
function _ChampExiste(NomTable, NomChamp : string) : Boolean;
{---------------------------------------------------------------------------------------}
BEGIN
{$IFDEF EAGLCLIENT}
  try
    Result := True;
    ExecuteSQL('SELECT '+NomChamp+' FROM '+NomTable);
  except
    on E: Exception do begin Result := False; Exit; end;
  end;
{$ELSE}
  Result:=ChampPhysiqueExiste(NomTable,NomChamp) ;
{$ENDIF}
END ;

{---------------------------------------------------------------------------------------}
procedure AfficheLeSolde (T : THNumEdit; TD, TC : Double);
{---------------------------------------------------------------------------------------}
BEGIN
if ((TD=TC) or (T.Tag=3)) then
   BEGIN
   T.NumericType:=ntGeneral ;
   if T.Value<>TD-TC then T.Value:=TD-TC ;
   END else if Abs(TD)>=Abs(TC) then
   BEGIN
   T.Debit:=True ; T.NumericType:=ntDC ;
   if T.Value<>TD-TC then T.Value:=TD-TC ;
   END else if Abs(TD)<Abs(TC) then
   BEGIN
   T.Debit:=False ; T.NumericType:=ntDC ;
   if T.Value<>TC-TD then T.Value:=TC-TD ;
   END ;
END ;

{---------------------------------------------------------------------------------------}
function SQLPremierDernier(fb : TFichierbase; Prem : Boolean) : string;
{---------------------------------------------------------------------------------------}
Var Q       : TQuery ;
    SQL,Cpt,LaTable,Desc,where : String ;
begin
result:='' ;
If Prem Then Desc:=' ASC ' Else Desc:=' DESC ' ;
Case fb Of
  fbGene : BEGIN
           Cpt:='G_GENERAL' ; LaTable:='GENERAUX' ; Where:=' ' ;
           END ;
  fbAux  : BEGIN
           Cpt:='T_AUXILIAIRE' ; LaTable:='TIERS' ; Where:=' ' ;
           END ;
  fbJAL  : BEGIN
           Cpt:='J_JOURNAL' ; LaTable:='JOURNAL' ; Where:=' ' ;
           END ;
  fbAxe1..fbAxe5 : BEGIN
                   Cpt:='S_SECTION' ; LaTable:='SECTION' ; Where:=' WHERE S_AXE="'+'A'+IntToStr(ord(fb)+1)+'" ' ;
                   END ;
  fbBudgen : BEGIN
             Cpt:='BG_BUDGENE' ; LaTable:='BUDGENE' ; Where:=' ' ;
             END ;
  fbBudSec1..fbBudSec5 : BEGIN
             Cpt:='BS_BUDSECT' ; LaTable:='BUDSECT' ; Where:=' WHERE BS_AXE="'+'A'+IntToStr(ord(fb)-13)+'" ' ;
             END ;
  fbBudJal : BEGIN
             Cpt:='BJ_BUDJAL' ; LaTable:='BUDJAL' ; Where:=' ' ;
             END ;
  END ;
SQL:='SELECT '+Cpt+' FROM '+LaTable+where+' ORDER BY '+Cpt+desc ;
Q:=OpenSQL(SQL,TRUE) ;
if Not Q.Eof then
   BEGIN
   Case fb Of
     fbGene : Result:=Q.FindField('G_GENERAL').AsString ;
     fbAux  : Result:=Q.FindField('T_AUXILIAIRE').AsString ;
     fbAxe1..fbAxe5 : Result:=Q.FindField('S_SECTION').AsString ;
     fbJal  : Result:=Q.FindField('J_JOURNAL').AsString ;
     fbBudgen : Result:=Q.FindField('BG_BUDGENE').AsString ;
     fbBudSec1..fbBudSec5 : Result:=Q.FindField('BS_BUDSECT').AsString ;
     fbBudJal : Result:=Q.FindField('BJ_BUDJAL').AsString ;
     end ;
   END ;
Ferme(Q) ;
end ;

{---------------------------------------------------------------------------------------}
procedure PremierDernierRub(ZoomTable : TZoomTable; SynPlus : string; var Cpt1, Cpt2 : string);
{---------------------------------------------------------------------------------------}

      {------------------------------------------------------------------}
      function SQLPremierDernierRub(ZoomTable : TZoomTable ; SynPlus : String ; Prem : Boolean) : String ;
      {------------------------------------------------------------------}
      Var Q       : TQuery ;
          SQL,Cpt,LaTable,Desc,where : String ;
      begin
      result:='' ;
      If Prem Then Desc:=' ASC ' Else Desc:=' DESC ' ;
      Cpt:='RB_RUBRIQUE' ; LaTable:='RUBRIQUE' ; Where:='' ;
      Case ZoomTable of
        tzRubCPTA  : Where:=' WHERE (RB_NATRUB="CPT")' ;
        tzRubBUDG  : Where:=' WHERE RB_NATRUB="BUD" AND RB_FAMILLES LIKE "CBG%" ' ;
        tzRubBUDS  : Where:=' WHERE RB_NATRUB="BUD" AND RB_FAMILLES LIKE "CBS%" ' ;
        tzRubBUDSG : Where:=' WHERE RB_NATRUB="BUD" AND RB_FAMILLES LIKE "S/G%" ' ;
        tzRubBUDGS : Where:=' WHERE RB_NATRUB="BUD" AND RB_FAMILLES LIKE "G/S%" ' ;
        Else Exit ;
        END ;
      If (SynPlus<>'') Then
        BEGIN
        If ZoomTable=tzRubCPTA Then Where:=Where+'AND RB_FAMILLES like "%'+SynPlus+'%"'
                               Else Where:=Where+'AND RB_BUDJAL like "%'+SynPlus+'%"' ;
        END ;
      SQL:='SELECT '+Cpt+' FROM '+LaTable+where+' ORDER BY '+Cpt+desc ;
      Q:=OpenSQL(SQL,TRUE) ;
      if Not Q.Eof then Result:=Q.FindField('RB_RUBRIQUE').AsString ;
      Ferme(Q) ;
      end ;

begin
  Cpt1 := SQLPremierDernierRub(ZoomTable,SynPlus,TRUE) ;
  Cpt2 := SQLPremierDernierRub(ZoomTable,SynPlus,FALSE) ;
end;

{---------------------------------------------------------------------------------------}
procedure PremierDernier(fb : TFichierBase; var Cpt1, Cpt2 : string);
{---------------------------------------------------------------------------------------}
begin
  Cpt1 := SQLPremierDernier(fb, True);
  Cpt2 := SQLPremierDernier(fb, False);
end ;

{Regarde la présence d'une fiche immobilisation dans la base
{---------------------------------------------------------------------------------------}
function AuMoinsUneImmo : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ExisteSQL('SELECT I_IMMO FROM IMMO');
end;

{---------------------------------------------------------------------------------------}
function PieceSurFolio(Jal : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery ;
begin
  Result := False;
  Q := OpenSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_JOURNAL="' + Jal + '" ', True);
  if not Q.Eof then
    Result := (Q.Fields[0].AsString = 'BOR') or (Q.Fields[0].AsString = 'LIB');
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
function  EstConfidentiel ( Conf : String ) : boolean ;
{---------------------------------------------------------------------------------------}
begin
  Result:=False ;
  if V_PGI.Confidentiel='1' then Exit ;
  if ((Conf='0') or (Conf='') or (Conf='-')) then Exit ;
  Result:=True ;
end ;

{---------------------------------------------------------------------------------------}
function  EstSQLConfidentiel(StTable : string; Cpte : String17) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Result := False;
  if V_PGI.Confidentiel = '1' then Exit;

       if StTable = 'GENERAUX' then Q := OpenSQL('SELECT G_CONFIDENTIEL FROM GENERAUX WHERE G_GENERAL="'+Cpte+'"',True)
  else if StTable = 'TIERS'    then Q := OpenSQL('SELECT T_CONFIDENTIEL FROM TIERS WHERE T_AUXILIAIRE="'+Cpte+'"',True)
  else Exit;

  if not Q.EOF then Result := EstConfidentiel(Q.Fields[0].AsString)
               else Result := False;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
procedure CorrespToCodes(Plan : THValComboBox; var C1, C2 : TComboBox);
{---------------------------------------------------------------------------------------}
Var St : String ;
    Q1 : TQuery ;
BEGIN
C1.Clear ; C2.Clear ;
St:='Select CR_CORRESP, CR_LIBELLE From CORRESP Where CR_TYPE="'+Plan.Value+'" '
   +'Order By CR_TYPE, CR_CORRESP' ;
Q1:=OpenSQL(St,True) ; Q1.First ;
While Not Q1.EOF do
    BEGIN     {A voir si on ajoute les Codes en value et les libellés en Items ... }
    C1.Items.Add(Q1.Fields[0].AsString) ;
    C2.Items.Add(Q1.Fields[0].AsString) ;
    Q1.Next ;
    END ;
Ferme(Q1) ;
END ;

{---------------------------------------------------------------------------------------}
procedure CreerDeviseTenue(LaMonnaie : string);
{---------------------------------------------------------------------------------------}
{$IFNDEF NOVH}
Var Q : TQuery ;
{$ENDIF NOVH}
BEGIN
{$IFNDEF NOVH}
If LaMonnaie='F' Then
  BEGIN
  Q:=OPENSQL('SELECT * FROM DEVISE WHERE D_DEVISE="FRF" ',FALSE) ;
  If Q.Eof Then
    BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('D_DEVISE').AsString:='FRF' ;
    Q.FindField('D_LIBELLE').AsString:='Francs' ;
    Q.FindField('D_SYMBOLE').AsString:='F' ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.FindField('D_DECIMALE').AsInteger:=2 ;
    Q.FindField('D_QUOTITE').AsInteger:=1 ;
    Q.FindField('D_MONNAIEIN').AsString:='X' ;
    Q.FindField('D_FONGIBLE').AsString:='X' ;
    Q.FindField('D_PARITEEURO').AsFloat:=6.55957 ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.Post ;
    END Else
    BEGIN
    Q.Edit ;
    Q.FindField('D_MONNAIEIN').AsString:='X' ;
    Q.FindField('D_FONGIBLE').AsString:='X' ;
    Q.FindField('D_PARITEEURO').AsFloat:=6.55957 ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.Post ;
    END ;
  Ferme(Q) ;
  ExecuteSQL('DELETE FROM DEVISE WHERE D_DEVISE="EUR"') ;
  SetParamSoc('SO_TENUEEURO',FALSE) ; SetParamSoc('SO_TAUXEURO',6.55957) ; SetParamSoc('SO_DECVALEUR',2) ;
  SetParamSoc('SO_DECPRIX',2) ;
  SetParamSoc('SO_DEVISEPRINC','FRF') ;
  END Else If LaMonnaie='E' Then
  BEGIN
  Q:=OPENSQL('SELECT * FROM DEVISE WHERE D_DEVISE="FRF" ',FALSE) ;
  If Q.Eof Then
    BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('D_DEVISE').AsString:='FRF' ;
    Q.FindField('D_LIBELLE').AsString:='Francs' ;
    Q.FindField('D_SYMBOLE').AsString:='F' ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.FindField('D_DECIMALE').AsInteger:=2 ;
    Q.FindField('D_QUOTITE').AsInteger:=1 ;
    Q.FindField('D_MONNAIEIN').AsString:='X' ;
    Q.FindField('D_FONGIBLE').AsString:='X' ;
    Q.FindField('D_PARITEEURO').AsFloat:=6.55957 ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.Post ;
    END Else
    BEGIN
    Q.Edit ;
    Q.FindField('D_MONNAIEIN').AsString:='X' ;
    Q.FindField('D_FONGIBLE').AsString:='X' ;
    Q.FindField('D_PARITEEURO').AsFloat:=6.55957 ;
    Q.FindField('D_FERME').AsString:='-' ;
    Q.Post ;
    END ;
  Ferme(Q) ;
  Q:=OPENSQL('SELECT * FROM DEVISE WHERE D_DEVISE="EUR" ',FALSE) ;
  If Q.Eof Then
    BEGIN
    Q.Insert ; InitNEw(Q) ;
    Q.FindField('D_DEVISE').AsString:='EUR' ;
    Q.FindField('D_LIBELLE').AsString:='Euro' ;
    Q.FindField('D_SYMBOLE').AsString:='EUR' ;
    Q.FindField('D_CODEISO').AsString:='EUR' ;
    Q.FindField('D_DECIMALE').AsInteger:=2 ;
    Q.FindField('D_QUOTITE').AsFloat:=1 ;
    Q.FindField('D_MONNAIEIN').AsString:='-' ;
    Q.FindField('D_FONGIBLE').AsString:='-' ;
    Q.FindField('D_PARITEEURO').AsFloat:=1 ;
    Q.FindField('D_FERME').AsString:='-' ;
    If VH^.RecupSISCOPGI Then
      BEGIN
      Q.FindField('D_MAXDEBIT').AsFloat:=1 ;
      Q.FindField('D_MAXCREDIT').AsFloat:=1 ;
      END ;
    Q.Post ;
    END Else
    BEGIN
    Q.Edit ;
    Q.FindField('D_MONNAIEIN').AsString:='-' ;
    Q.FindField('D_FONGIBLE').AsString:='-' ;
    Q.FindField('D_PARITEEURO').AsFloat:=1 ;
    Q.FindField('D_FERME').AsString:='-' ;
    If VH^.RecupSISCOPGI Then
      BEGIN
      Q.FindField('D_MAXDEBIT').AsFloat:=1 ;
      Q.FindField('D_MAXCREDIT').AsFloat:=1 ;
      END ;
    Q.Post ;
    END ;
  Ferme(Q) ;
  SetParamSoc('SO_DEVISEPRINC','EUR') ;
  SetParamSoc('SO_TENUEEURO',TRUE) ; SetParamSoc('SO_TAUXEURO',1) ; SetParamSoc('SO_DECVALEUR',2) ;
  SetParamSoc('SO_DECPRIX',2) ;
  END ;
{$ENDIF NOVH}
END ;

function CleSequence(TypeSouche,CodeSouche:string;DD:tDateTime;MulTiExo:boolean) : string;
begin
  if MultiExo then
    Result:=Cop(TypeSouche,3)+Cop(CodeSouche,3)+cop(CodeExo(DD),3)
  else Result:=Cop(TypeSouche,3)+Cop(CodeSouche,3);
end;

function ExistSequence(Cle: string) : Boolean;
begin
	ExisteQl
end;

function CreerCompteur(TypeSouche,CodeSouche:string;DD:tDateTime;Compteur:integer;MultiExo:boolean;Entity:integer=0):integer;
var Cle : string;
begin
  if Entity=0 then Entity:=Entite;
  Cle:=CleSequence(TypeSouche,CodeSouche,DD,MultiExo);
  if not ExistSequence(Cle) then
    CreateSequence(Cle,Compteur,1);
  Result:=ReadCurrent(Cle);
end;

function VerifCaractereInterditSouche : string ;
Var ST,NewSt,SauveSt,SH_TYPE : String ;
    Q : tquery ;
    TobS,TobL : tob ;
    TobSAll : tob ;
    i : Integer ;
    Car : Char ;
    OkOk :Boolean ;
begin
  Result:='' ;
  TobS := TOB.Create('', nil, -1);
  TobSAll := TOB.Create('', nil, -1);
  try
    Try
    St:='SELECT * FROM SOUCHE WHERE SH_TYPE IN ("CPT","BUD","REL","TRE") AND SH_SOUCHE LIKE "%.%" ' ;
    Q:=OpenSQL(St,True) ;
    TobS.LoadDetailDB ('SOUCHE', '', '', Q, True);
    Ferme(Q) ;

    St:='SELECT * FROM SOUCHE WHERE SH_TYPE IN ("CPT","BUD","REL","TRE") ' ;
    Q:=OpenSQL(St,True) ;
    TobSAll.LoadDetailDB ('SOUCHE', '', '', Q, True);
    Ferme(Q) ;

    If TobS.Detail.Count>0 then
      begin
      For i:=0 To TobS.Detail.Count-1 Do
        BEGIN

        TobL:=TobS.Detail[i] ;
        SH_TYPE:=TobL.GetValue('SH_TYPE') ;
        St:=TobL.GetValue('SH_SOUCHE') ; SauveSt:=St ;
        Car:='A' ;
        Repeat
        NewSt:=FindEtReplace(St, '.', Car, True);
        OkOk:=TRUE ;

        If TobSAll.FindFirst(['SH_TYPE','SH_SOUCHE'],[SH_TYPE,NewSt],TRUE) <>NIL Then
           begin
           OkOK:=False ;
           If car='Z' then BEGIN OkOk:=TRUE ; Result:=Result+SH_TYPE+';'+SauveSt+';' ; END  Else Car:=Succ(Car) ;
           END ;
        If Not OkOk Then St:=SauveSt ;
        until OkOk ;
        If OkOk then
          begin
          If SH_TYPE='CPT' then
            begin
            ExecuteSQL('UPDATE JOURNAL SET J_COMPTEURNORMAL="'+NewSt+'" WHERE J_COMPTEURNORMAL="'+SauveSt+'" ') ;
            ExecuteSQL('UPDATE JOURNAL SET J_COMPTEURSIMUL="'+NewSt+'" WHERE J_COMPTEURSIMUL="'+SauveSt+'" ');
            end Else
          If SH_TYPE='BUD' then
            begin
            ExecuteSQL('UPDATE BUDJAL SET BJ_COMPTEURNORMAL="'+NewSt+'" WHERE BJ_COMPTEURNORMAL="'+St+'" ');
            ExecuteSQL('UPDATE BUDJAL SET BJ_COMPTEURSIMUL="'+NewSt+'" WHERE BJ_COMPTEURSIMUL="'+St+'" ');
            end ;
          ExecuteSQL('UPDATE SOUCHE SET SH_SOUCHE="'+NewSt+'" WHERE SH_TYPE="'+SH_TYPE+'" AND SH_SOUCHE="'+SauveSt+'" ')
          end;
        END ;
      END ;
    except
      Result:='ERREUR' ;
      raise Exception.Create(traduirememoire('Création Séquence : Problème de transformation des codes souches (Caractère ".")'));
    End ;
  Finally
    TobS.ClearDetail ; TobS.Free ;
    TobSAll.ClearDetail ; TobSAll.Free ;
  end ;
END ;

Procedure CPInitSequenceCompta;
//Lek 170409 Créer les séquences compta selon table SOUCHE de toutes les entités
var lQ : tQuery;
    req,St : String;
begin
St:=VerifCaractereInterditSouche ;
If St='ERREUR' Then Exit ;
  req:='SELECT * FROM SOUCHE WHERE SH_TYPE IN ("CPT","BUD","REL","TRE")';
  if OkMultiEntite then
    req:='@NOENTITY@ '+req;
  try
    lQ:=OpenSql(req,TRUE);
    while not lQ.Eof do begin
      If OkSouche(lQ,St) then
        BEGIN
        if lQ.FindField('SH_SOUCHEEXO').AsString='X' then begin
          if lQ.FindField('SH_NUMDEPARTS').AsInteger >0 then
            CreerCompteur(lQ.FindField('SH_TYPE').AsString,
                          lQ.FindField('SH_SOUCHE').AsString,
                          GetSuivant.Deb,
                          lQ.FindField('SH_NUMDEPARTS').AsInteger,
                          lQ.FindField('SH_SOUCHEEXO').AsString='X',
                          lQ.FindField('SH_ENTITY').AsInteger);
          if lQ.FindField('SH_NUMDEPARTP').AsInteger >0 then
            CreerCompteur(lQ.FindField('SH_TYPE').AsString,
                          lQ.FindField('SH_SOUCHE').AsString,
                          GetPrecedent.Deb,
                          lQ.FindField('SH_NUMDEPARTP').AsInteger,
                          lQ.FindField('SH_SOUCHEEXO').AsString='X',
                          lQ.FindField('SH_ENTITY').AsInteger);
        end;
        if lQ.FindField('SH_NUMDEPART').AsInteger >0 then
          CreerCompteur(lQ.FindField('SH_TYPE').AsString,
                        lQ.FindField('SH_SOUCHE').AsString,
                        GetEncours.Deb,
                        lQ.FindField('SH_NUMDEPART').AsInteger,
                        lQ.FindField('SH_SOUCHEEXO').AsString='X',
                        lQ.FindField('SH_ENTITY').AsInteger);
      END ;
      lQ.Next;
    end;
  finally Ferme(lQ); end;
end;


end.
