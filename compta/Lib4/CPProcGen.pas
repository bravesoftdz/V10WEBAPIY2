unit CPProcGen;

interface

uses
  Hctrls, Forms, HTB97, Buttons, Menus, Classes, Controls, Dialogs, uTob, HEnt1;


procedure ChangeMask(C : THNumEdit; Decim : Integer; Symbole : string);
function  EstJoker(St : string): Boolean;

{Gestion des dates et des périodes}
function  IncPeriode(Periode : Integer) : Integer;
function  GetPeriode( LaDate : TDateTime ) : Integer;
function  GetDateTimeFromPeriode(Periode : Integer) : TDateTime;
function  CIsValideDate(aDate : string) : Boolean;
function  CConvertDate (DateChar : string) : TDateTime; {aDate au format JJMMAA}
function  ConvertMontantCFONB(MontantChar : string; Deci : Integer) : Double;

{Gestion des formats de montant}
function  PrintSoldeFormate(TD, TC : Double; Decim : Integer; Symbole : string; AffSymb : Boolean; FormMont : string) : string;
function  PrintEcart(TD, TC : Double; Decim : Integer; DebitPos : Boolean) : string;
function  PrintSolde(TD, TC : Double; Decim : Integer; Symbole : string; AffSymb : Boolean) : string;
function  AfficheMontant(Formatage, LeSymbole : string; LeMontant : Double; OkSymbole : Boolean) : string;
function  TotDifferent(X1, X2 : Double) : Boolean;
function  MoinsSymbole(LeSolde : string) : Double;

{Gestion de Chaînes}
function  OkChar(St : string) : Boolean;
function  Fourchette(var St : string) : string;
function  Insere(st, st1 : string; Deb, Long : Integer) : string;
{Supprime les accents des voyelles et met les voyelles en majuscule}
procedure ModifCarAccentue(var St : string);
{Supprime les accents des voyelles mais laisse les voyelles dans leur casse}
procedure ModifCarAccentueKeep(var St : string);
{Transforme les voyelles accentuées de minuscule en majuscule}
procedure ModifCarAccentueMajMin(var St : string);
{ Retourne la longueur en pixel à partir d'un nombre de caractères }
function  GetPixelFromCar ( Ecran : TForm ; vInCar : integer ) : integer ;
function  GetCarFromPixel ( Ecran : TForm ; vInPixel : integer ) : integer ;


{Gestion de fichiers}
procedure DirDefault(Sauve : TSaveDialog ; FileName : string);
function  FileTemp(const aExt : string) : string;
{22/10/07 : FQ 21706 : transforme un fichier sans retour de chariot en fichier avec}
function TransformeFichier(NomFichier : string; lgEnreg : Integer = 120) : string;
{Decoupe un fichier (CFONB) contenant plusieurs "ensembles" :
 Deb/Fin = enregistrement de début et de fin 01/07 ou 31/39 ...}
procedure DecoupeFichier(NomFichier, Deb, Fin, Ext : string);

{Gestion bancaire}
function QuelPaysLocalisation : string;
function CalcRIB(Pays : string; Banque : string; Guichet : string; Compte : string; Cle : string) : string;
function calcIBAN(Pays : string; RIB : string) : string;

{$IFNDEF EAGLSERVER}
{Gestion des PopupMenus}
procedure PopZoom97(BM : TToolbarButton97; POPZ : TPopupMenu);
procedure PurgePopup(PP : TPopupMenu);
procedure PopZoom(BM : TBitBtn; POPZ : TPopupMenu);
function  IsButtonPop (CC : TComponent) : Boolean;
function  CreerLigPop(B : TBitBtn; Owner : TComponent; ShortC : Boolean; C : Char) : TMenuItem;
function  AttribShortCut(Nom : string) : string;
procedure InitPopup(F : TForm);
{$ENDIF EAGLSERVER}

{Divers en rapport avec l'AGL}
function  TAToStr(TA : TActionFiche) : string;
function  StrToTA(psz : string) : TActionFiche;
procedure AvertirMultiTable(FTypeTable : string);

// GCO - 20/10/2008
function  CreateTobParamCWA( vStNameAppli : HString ) : Tob;

implementation

uses
  UtilPgi, {Fonctions de gestion des RIB et des IBAN}
  SysUtils, ComCtrls, Windows, ParamSoc, Messages, CPTypeCons;


{---------------------------------------------------------------------------------------}
function GetPeriode ( LaDate : TDateTime ) : Integer;
{---------------------------------------------------------------------------------------}
var
  YY,MM,DD : Word;
begin
  DecodeDate(LaDate, YY, MM, DD);
  Result := 100 * YY + MM;
end;

{---------------------------------------------------------------------------------------}
function IncPeriode(Periode : Integer) : Integer;
{---------------------------------------------------------------------------------------}
begin
  if System.Copy(IntToStr(Periode), 5, 2) = '12' then
    Result := Periode + 89
  else
    Result := Periode + 1;
end;

{---------------------------------------------------------------------------------------}
function  GetDateTimeFromPeriode(Periode : Integer) : TDateTime;
{---------------------------------------------------------------------------------------}
begin
  Result := EncodeDate(StrToInt(Copy(IntToStr(Periode), 1, 4)), StrToInt(Copy(IntToStr(Periode), 5, 2)), 1);
end;

{Fonction un peu plus exigeante que IsValideDate et qui teste si la date en
 paramêtre est comprise entre 01/01/1900 et 31/12/2099
{---------------------------------------------------------------------------------------}
function CIsValideDate(aDate : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := IsValidDate(aDate);
  if Result then
    Result := (iDate1900 <= StrToDate(aDate)) and (iDate2099 >= StrToDate(aDate));
end;

{---------------------------------------------------------------------------------------}
function CConvertDate(DateChar : string) : TDateTime; {DateChar au format JJMMAA}
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin
  if (DateChar = '000000') or (Trim(DateChar) = '') then
    Result := iDate1900

  else begin
    a  := StrToInt(Copy(DateChar, 5, 2));
    if a > 90 then a := a + 1900
              else a := a + 2000;
    m := StrToInt(Copy(DateChar, 3, 2));
    j := StrToInt(Copy(DateChar, 1, 2));
    Result:= EncodeDate(a, m, j);
  end;
end;

{---------------------------------------------------------------------------------------}
function ConvertMontantCFONB(MontantChar : string; Deci : Integer) : Double;
{---------------------------------------------------------------------------------------}
var
  MontantFloat : Double;
  Lettre : Char;
  Divi   : Integer;
  n      : Integer;
  lInSize: integer ;
begin
(*
A=+1 , B=+2 , C=+3 , D=+4 , E=+5 , F=+6 , G=+7 , H=+8 , I=+9
 J=-1 , K=-2 , L=-3 , M=-4 , N=-5 , 0=-6 , P=-7 , Q=-8 , R=-9
é=+0 ( ou {=+0)
é=-0 ( ou }=-0)
*)
  // SBO 01/04/2008 : Plantage systématique car ce sont des chaînes de 12 caractères qui sont passées en paramètre dans le cadre de l'import CFONB
  lInSize := Length( MontantChar ) ;
  Lettre := MontantChar[lInSize] ;

  Divi := 1;
  MontantFloat :=  1;

  for n := 1 to Deci do Divi := Divi * 10;

  case Lettre of
    '{'      : Lettre := '0';
    '}'      : begin
                 MontantFloat := -1;
                 Lettre := '0';
               end;
    'A'..'I' : Dec(Lettre, 16);
    'J'..'R' : begin
                 MontantFloat := -1;
                 Dec(Lettre, 25);
               end;
  end;

  Result := MontantFloat * Valeur(Copy(MontantChar, 1, (lInSize-1) ) + Lettre) / Divi;

end;

{---------------------------------------------------------------------------------------}
procedure ChangeMask(C : THNumEdit; Decim : Integer; Symbole : string);
{---------------------------------------------------------------------------------------}
var
  j : Integer;
begin
  C.Decimals := Decim;
  C.CurrencySymbol:= Trim(Symbole);
  if Decim = 0 then C.Masks.PositiveMask := '#,##0'
               else C.Masks.PositiveMask := '#,##0.';

  for j := 1 to Decim do
    C.Masks.PositiveMask := C.Masks.PositiveMask + '0';

  C.Masks.ZeroMask := C.Masks.PositiveMask;
  C.Masks.NegativeMask := '-' + C.Masks.PositiveMask;
end;

{---------------------------------------------------------------------------------------}
function EstJoker(St : string): Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
  if (pos('%', St) > 0)
    or (pos('[', St) > 0)
      or (pos('_', St) > 0)
        or (pos(']', St) > 0)
          or (pos('?', St) > 0)
            or (pos('*', St) > 0) then Result := True;
end;

{---------------------------------------------------------------------------------------}
function FileTemp(const aExt : string) : string;
{---------------------------------------------------------------------------------------}
var
  Buffer: array[0..1023] of Char;
  aFile : string;
begin
  GetTempPath(Sizeof(Buffer) - 1, Buffer);
  GetTempFileName(Buffer, 'TMP', 0, Buffer);
  SetString(aFile, Buffer, StrLen(Buffer));
  Result := ChangeFileExt(aFile, aExt);
  RenameFile(aFile, Result);
end;

{---------------------------------------------------------------------------------------}
procedure DirDefault(Sauve : TSaveDialog ; FileName : string);
{---------------------------------------------------------------------------------------}
var
  j, i : Integer;
begin
  j := Length(FileName);
  for i := Length(FileName) downto 1 do
    if FileName[i] = '\' then begin
      j := i;
      Break;
    end;

  Sauve.InitialDir := Copy(FileName, 1, j);
end;

{---------------------------------------------------------------------------------------}
function Fourchette(var St : string) : string;
{---------------------------------------------------------------------------------------}
var
  i : Integer ;
begin
  i := Pos(':', St);
  if i <= 0 then i := Length(St) + 1;
  Result := Copy(St, 1, i - 1);
  Delete(St, 1, i);
end;

{---------------------------------------------------------------------------------------}
function OkChar(St : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  i : Integer;
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','OkChar') ;
 {$ENDIF}
  Result := false ;
  for i := 1 to Length(St) do
    if not (St[i] in ['0'..'9']) then Exit;

  Result := True;
end;

{---------------------------------------------------------------------------------------}
function Insere(st, st1 : string; Deb, Long : Integer) : string;
{---------------------------------------------------------------------------------------}
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.ChargeMagExo') ;
{$ENDIF}
  Insere := Copy(st, 1, Deb - 1) +
            Copy(st1, 1, Long) +
            Copy(st, Deb + Long, length(st) - Long);
end;

{Supprime les accents des voyelles et met les voyelles en majuscule
{---------------------------------------------------------------------------------------}
procedure ModifCarAccentue(var St : string);
{---------------------------------------------------------------------------------------}
var
  p : integer;
begin
  // fiche 20949
  repeat begin p := Pos('É',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('È',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('Ê',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('Ë',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('é',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('è',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('ê',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('ë',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('é',St); if p >0 then St[p] := 'E'; end until not (p > 0);


  repeat begin p := Pos('À',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Á',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Â',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Ã',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Ä',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Å',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Æ',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('ä',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('â',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('à',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('á',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('â',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('ã',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('ä',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('å',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('æ',St); if p >0 then St[p] := 'A'; end until not (p > 0);

  repeat begin p := Pos('ç',St); if p >0 then St[p] := 'C'; end until not (p > 0);
  repeat begin p := Pos('Ç',St); if p >0 then St[p] := 'C'; end until not (p > 0);

  repeat begin p := Pos('ù',St); if p >0 then St[p] := 'U'; end until not (p > 0);
  repeat begin p := Pos('ü',St); if p >0 then St[p] := 'U'; end until not (p > 0);
  repeat begin p := Pos('û',St); if p >0 then St[p] := 'U'; end until not (p > 0);
  repeat begin p := Pos('ú',St); if p >0 then St[p] := 'U'; end until not (p > 0);
  repeat begin p := Pos('Û',St); if p >0 then St[p] := 'U'; end until not (p > 0);
  repeat begin p := Pos('Ù',St); if p >0 then St[p] := 'U'; end until not (p > 0);
  repeat begin p := Pos('Ú',St); if p >0 then St[p] := 'U'; end until not (p > 0);
  repeat begin p := Pos('Ü',St); if p >0 then St[p] := 'U'; end until not (p > 0);

  repeat begin p := Pos('î',St); if p >0 then St[p] := 'I'; end until not (p > 0);
  repeat begin p := Pos('ï',St); if p >0 then St[p] := 'I'; end until not (p > 0);
  repeat begin p := Pos('ì',St); if p >0 then St[p] := 'I'; end until not (p > 0);
  repeat begin p := Pos('í',St); if p >0 then St[p] := 'I'; end until not (p > 0);
  repeat begin p := Pos('Î',St); if p >0 then St[p] := 'I'; end until not (p > 0);
  repeat begin p := Pos('Ï',St); if p >0 then St[p] := 'I'; end until not (p > 0);
  repeat begin p := Pos('Ì',St); if p >0 then St[p] := 'I'; end until not (p > 0);
  repeat begin p := Pos('Í',St); if p >0 then St[p] := 'I'; end until not (p > 0);

  repeat begin p := Pos('ô',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('ö',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('ò',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('ó',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('õ',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('Ò',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('Ó',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('Ô',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('Õ',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('Ö',St); if p >0 then St[p] := 'O'; end until not (p > 0);
end;

{Supprime les accents des voyelles mais laisse les voyelles dans leur casse
{---------------------------------------------------------------------------------------}
procedure ModifCarAccentueKeep(var St : string);
{---------------------------------------------------------------------------------------}
var
  p : Integer;
begin
  repeat begin p := Pos('É',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('È',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('Ê',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('Ë',St); if p >0 then St[p] := 'E'; end until not (p > 0);
  repeat begin p := Pos('è',St); if p >0 then St[p] := 'e'; end until not (p > 0);
  repeat begin p := Pos('ê',St); if p >0 then St[p] := 'e'; end until not (p > 0);
  repeat begin p := Pos('ë',St); if p >0 then St[p] := 'e'; end until not (p > 0);
  repeat begin p := Pos('é',St); if p >0 then St[p] := 'e'; end until not (p > 0);


  repeat begin p := Pos('À',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Á',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Â',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Ã',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Ä',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Å',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('Æ',St); if p >0 then St[p] := 'A'; end until not (p > 0);
  repeat begin p := Pos('ä',St); if p >0 then St[p] := 'a'; end until not (p > 0);
  repeat begin p := Pos('â',St); if p >0 then St[p] := 'a'; end until not (p > 0);
  repeat begin p := Pos('à',St); if p >0 then St[p] := 'a'; end until not (p > 0);
  repeat begin p := Pos('á',St); if p >0 then St[p] := 'a'; end until not (p > 0);
  repeat begin p := Pos('â',St); if p >0 then St[p] := 'a'; end until not (p > 0);
  repeat begin p := Pos('ã',St); if p >0 then St[p] := 'a'; end until not (p > 0);
  repeat begin p := Pos('ä',St); if p >0 then St[p] := 'a'; end until not (p > 0);
  repeat begin p := Pos('å',St); if p >0 then St[p] := 'a'; end until not (p > 0);
  repeat begin p := Pos('æ',St); if p >0 then St[p] := 'a'; end until not (p > 0);

  repeat begin p := Pos('ç',St); if p >0 then St[p] := 'c'; end until not (p > 0);
  repeat begin p := Pos('Ç',St); if p >0 then St[p] := 'C'; end until not (p > 0);

  repeat begin p := Pos('ù',St); if p >0 then St[p] := 'u'; end until not (p > 0);
  repeat begin p := Pos('ü',St); if p >0 then St[p] := 'u'; end until not (p > 0);
  repeat begin p := Pos('û',St); if p >0 then St[p] := 'u'; end until not (p > 0);
  repeat begin p := Pos('ú',St); if p >0 then St[p] := 'u'; end until not (p > 0);
  repeat begin p := Pos('Û',St); if p >0 then St[p] := 'U'; end until not (p > 0);
  repeat begin p := Pos('Ù',St); if p >0 then St[p] := 'U'; end until not (p > 0);
  repeat begin p := Pos('Ú',St); if p >0 then St[p] := 'U'; end until not (p > 0);
  repeat begin p := Pos('Ü',St); if p >0 then St[p] := 'U'; end until not (p > 0);

  repeat begin p := Pos('î',St); if p >0 then St[p] := 'i'; end until not (p > 0);
  repeat begin p := Pos('ï',St); if p >0 then St[p] := 'i'; end until not (p > 0);
  repeat begin p := Pos('ì',St); if p >0 then St[p] := 'i'; end until not (p > 0);
  repeat begin p := Pos('í',St); if p >0 then St[p] := 'i'; end until not (p > 0);
  repeat begin p := Pos('Î',St); if p >0 then St[p] := 'I'; end until not (p > 0);
  repeat begin p := Pos('Ï',St); if p >0 then St[p] := 'I'; end until not (p > 0);
  repeat begin p := Pos('Ì',St); if p >0 then St[p] := 'I'; end until not (p > 0);
  repeat begin p := Pos('Í',St); if p >0 then St[p] := 'I'; end until not (p > 0);

  repeat begin p := Pos('ô',St); if p >0 then St[p] := 'o'; end until not (p > 0);
  repeat begin p := Pos('ö',St); if p >0 then St[p] := 'o'; end until not (p > 0);
  repeat begin p := Pos('ò',St); if p >0 then St[p] := 'o'; end until not (p > 0);
  repeat begin p := Pos('ó',St); if p >0 then St[p] := 'o'; end until not (p > 0);
  repeat begin p := Pos('õ',St); if p >0 then St[p] := 'o'; end until not (p > 0);
  repeat begin p := Pos('Ò',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('Ó',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('Ô',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('Õ',St); if p >0 then St[p] := 'O'; end until not (p > 0);
  repeat begin p := Pos('Ö',St); if p >0 then St[p] := 'O'; end until not (p > 0);
end;

{Transforme les voyelles accentuées de minuscule en majuscule
{---------------------------------------------------------------------------------------}
procedure ModifCarAccentueMajMin(var St : string);
{---------------------------------------------------------------------------------------}
var
  p : integer;
begin
  repeat begin p := Pos('É',St); if p >0 then St[p] := 'é'; end until not (p > 0);
  repeat begin p := Pos('È',St); if p >0 then St[p] := 'è'; end until not (p > 0);
  repeat begin p := Pos('Ê',St); if p >0 then St[p] := 'ê'; end until not (p > 0);
  repeat begin p := Pos('Ë',St); if p >0 then St[p] := 'ë'; end until not (p > 0);


  repeat begin p := Pos('À',St); if p >0 then St[p] := 'à'; end until not (p > 0);
  repeat begin p := Pos('Á',St); if p >0 then St[p] := 'á'; end until not (p > 0);
  repeat begin p := Pos('Â',St); if p >0 then St[p] := 'â'; end until not (p > 0);
  repeat begin p := Pos('Ã',St); if p >0 then St[p] := 'â'; end until not (p > 0);
  repeat begin p := Pos('Ä',St); if p >0 then St[p] := 'ä'; end until not (p > 0);
  repeat begin p := Pos('Å',St); if p >0 then St[p] := 'å'; end until not (p > 0);
  repeat begin p := Pos('Æ',St); if p >0 then St[p] := 'æ'; end until not (p > 0);


  repeat begin p := Pos('Ç',St); if p >0 then St[p] := 'ç'; end until not (p > 0);

  repeat begin p := Pos('Û',St); if p >0 then St[p] := 'û'; end until not (p > 0);
  repeat begin p := Pos('Ù',St); if p >0 then St[p] := 'ù'; end until not (p > 0);
  repeat begin p := Pos('Ú',St); if p >0 then St[p] := 'ú'; end until not (p > 0);
  repeat begin p := Pos('Ü',St); if p >0 then St[p] := 'ü'; end until not (p > 0);

  repeat begin p := Pos('Î',St); if p >0 then St[p] := 'î'; end until not (p > 0);
  repeat begin p := Pos('Ï',St); if p >0 then St[p] := 'ï'; end until not (p > 0);
  repeat begin p := Pos('Ì',St); if p >0 then St[p] := 'ì'; end until not (p > 0);
  repeat begin p := Pos('Í',St); if p >0 then St[p] := 'í'; end until not (p > 0);

  repeat begin p := Pos('Ò',St); if p >0 then St[p] := 'ò'; end until not (p > 0);
  repeat begin p := Pos('Ó',St); if p >0 then St[p] := 'ó'; end until not (p > 0);
  repeat begin p := Pos('Ô',St); if p >0 then St[p] := 'ô'; end until not (p > 0);
  repeat begin p := Pos('Õ',St); if p >0 then St[p] := 'õ'; end until not (p > 0);
  repeat begin p := Pos('Ö',St); if p >0 then St[p] := 'ö'; end until not (p > 0);
end;

function  GetPixelFromCar( Ecran : TForm ; vInCar : integer ) : integer ;
begin
  Result := TForm(Ecran).Canvas.TextWidth(StringOfChar('W',vInCar)) ;
end ;
function  GetCarFromPixel ( Ecran : TForm ; vInPixel : integer ) : integer ;
begin
  Result := Round(vInPixel / TForm(Ecran).Canvas.TextWidth(StringOfChar('W',1))) ;
end;

{---------------------------------------------------------------------------------------}
function QuelPaysLocalisation : string;
{---------------------------------------------------------------------------------------}
begin
  Result := GetParamSocSecur('SO_PAYSLOCALISATION', '');
  if Result = '' then
    Result := CodeISOFR;
end;

{---------------------------------------------------------------------------------------}
function CalcRIB(Pays : string; Banque : string; Guichet : string; Compte : string; Cle : string) : string;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF COMPTA}
  if QuelPaysLocalisation=CodeISOES then
    Result:=FindetReplace(FindetReplace(EncodeRib(Banque,Guichet,Compte,Cle,'',Pays),' ','',TRUE),'/','',TRUE)
  else Begin
  {$ENDIF}
    if (Pays = CodeISOES) then
      Result := Trim(banque) + Trim(guichet) + Trim(cle) + Trim(compte)
    else if (pays = CodeISOFR) then
      Result := Trim(banque) + Trim(guichet) + Trim(compte) + Trim(cle)
    else
      Result := Trim(banque) + Trim(guichet) + Trim(compte) + Trim(cle);
  {$IFDEF COMPTA}
   end;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function calcIBAN(Pays : string; RIB : string) : string;
{---------------------------------------------------------------------------------------}
var
  St2, St4,
  cleIBAN,
  ret, strInter : string;
		ii   : Byte;
		cleL : Integer;
		i    : Integer;
begin
		Result := '';
		St2 := Trim(RIB) + Trim(Pays) + '00';

		if Length(St2) < 10 then Exit;

		St2 := UpperCase(St2);

		{Transforme les lettres en chiffres selon le NEE 5.3}
		i := 1;
		while i < Length(St2) do begin
 			if St2[i] in ['A'..'Z'] then	begin
	 			 ii  := Ord(St2[i]) - 65;
		 	 	st4 := Copy(st2, 1, i - 1) + IntToStr(10 + ii) + Copy(st2, i + 1, Length(st2));
			  	st2 := st4;
			 end;
			 Inc(i);
		end;

		ret  := '';
		cleL := 0;
		st4  := '';
		{On découpe par tranche de 9
		 On calcul la clé via mod 97 puis on fait clé + reste du rib
   XMG 14/07/03 début (depuis le deuxième passage on prennait CleL+9 caractères de St2, ce qui n'est pas ce qu'on attendaît)....}
  if QuelPaysLocalisation = CodeISOES then Begin
     while length(st2) > 0 do begin
       st2:= FormatFloat('####',cleL)+St2 ;
       st4 := Copy(st2, 1, 9) ;
       Delete(st2, 1, 9);
       cleL := StrToInt64(st4) mod 97;
     end;
  end
  else Begin
    for i:=1 to (length(st2) div 9) + 1 do begin
      st4 := Copy(st2, 1, 9) ;
      Delete(st2, 1, 9);
      StrInter := IntToStr(cleL) + st4;
      cleL := StrToInt64(StrInter) mod 97;
    end;
  end;

		{Une fois fini, on calcul 98-clé}
		cleIBAN := IntToStr(98 - (cleL  mod 97));
		if Length(cleIBAN)=1 then cleIBAN := '0' + cleIBAN;
		Result:= Trim(Pays)+ Trim(cleIBAN) + Trim(RIB);
end ;



{$IFNDEF EAGLSERVER}
{---------------------------------------------------------------------------------------}
procedure PurgePopup(PP : TPopupMenu);
{---------------------------------------------------------------------------------------}
var
  M, N : TMenuItem ;
begin
  if PP = nil then Exit;
  if PP.Items.Count <= 0 then Exit;
  while PP.Items.Count>0 do begin
    M := PP.Items[0] ;
    while M.Count>0 do begin
      N := M.Items[0];
      M.Remove(N);
      N.Free;
    end;
    PP.Items.Remove(M);
    M.Free;
  end;
end;

{---------------------------------------------------------------------------------------}
function IsButtonPop (CC : TComponent) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ((CC.ClassType = TBitBtn) or
             (CC.ClassType = THBitBtn) or
             (CC.ClassType = TSpeedButton) or
             (CC.ClassType = TToolbarButton97));
end ;

{---------------------------------------------------------------------------------------}
function CreerLigPop(B : TBitBtn; Owner : TComponent; ShortC : Boolean; C : Char) : TMenuItem;
{---------------------------------------------------------------------------------------}
var
  Nom, SC : string;
  T       : TMenuItem;
begin
  Nom := B.Name;
  System.Delete(Nom, 1, 1);
  T := TMenuItem.Create(Owner);
  T.Name := C + Nom;
  T.Caption := B.Hint;
  T.OnClick := B.OnClick;
  T.Tag := B.Tag;
  if ShortC then SC := AttribShortCut(Nom)
            else SC := '';
  if SC <> '' then T.Caption := T.Caption + #9 + SC;
  Result := T;
end ;


{---------------------------------------------------------------------------------------}
procedure PopZoom(BM : TBitBtn; POPZ : TPopupMenu);
{---------------------------------------------------------------------------------------}
var
  i : integer;
  T : TMenuItem;
  B : TBitBtn;
  P : TPoint;
  F : TForm;
  LeMsg : TMsg;
  OkVisible : Boolean;
  TTB : TTabSheet;
begin
  PurgePopup(POPZ);
  F := TForm(BM.Owner);

  for i := 0 to F.ComponentCount - 1 do begin
    if (IsButtonPop(F.Components[i])) then
      if ((F.Components[i].Tag = -BM.Tag) and (TControl(F.Components[i]).Parent.Enabled)) then begin
        if (TControl(F.Components[i]).Parent is ttabSheet) then begin
          TTB := TTabSheet(TControl(F.Components[i]).Parent);
          OkVisible := TTB.TabVisible;
        end
        else
          OkVisible := TControl(F.Components[i]).Parent.Visible;

        if OkVisible then begin
          B := TBitBtn(F.Components[i]);
          if (B.Enabled) then begin
            T := CreerLigPop(B, POPZ, False, 'Z');
            POPZ.Items.Add(T);
          end;
        end;
      end;
  end;

  P.X := 0;
  P.Y := BM.Height;
  P := BM.ClientToScreen(P);
  POPZ.PopUp(P.X,P.Y);

  while PeekMessage(LeMsg,HWND(0), WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE) do ;
end;

{BPY 26/02/2003
 - modification du test pour savoir si le Component est affecté ....
   car en eAGL on as des Components sans parent donc plantage ....
 - ajout aussi de la fin de la fonction ttabSheet ... maintenant
   elle afiche un menu ;)
 - de plus j'ai recopié le test supplementaire du ttabSheet qui
   n'existait pas dans ttabSheet ...
{---------------------------------------------------------------------------------------}
procedure PopZoom97(BM : TToolbarButton97; POPZ : TPopupMenu);
{---------------------------------------------------------------------------------------}
var
  i : integer;
  T : TMenuItem;
  B : TBitBtn;
  F : TForm;
  OkVisible : Boolean;
  TTB : TTabSheet;
begin
  PurgePopup(POPZ);
  F := TForm(BM.Owner);

  for i := 0 to F.ComponentCount - 1 do begin
    if (IsButtonPop(F.Components[i])) then
      if ((F.Components[i].Tag = -BM.Tag) and
          (TControl(F.Components[i]).Parent.Enabled) and
          (TControl(F.Components[i]).Parent.Visible)) then begin
        if (TControl(F.Components[i]).Parent is ttabSheet) then begin
          TTB := TTabSheet(TControl(F.Components[i]).Parent);
          OkVisible := TTB.TabVisible;
        end
        else
          OkVisible := TControl(F.Components[i]).Parent.Visible ;

        if OkVisible then begin
          B := TBitBtn(F.Components[i]) ;
          if (B.Enabled) then begin
            T := CreerLigPop(B,POPZ,False,'Z');
            POPZ.Items.Add(T);
          end;
        end;
      end;
  end;
end;

{---------------------------------------------------------------------------------------}
function AttribShortCut(Nom : string) : string;
{---------------------------------------------------------------------------------------}
var
  RS : TShortCut;
begin
  Result:='' ; RS:=0 ;
  Nom:=uppercase(Nom) ;
  //if Nom='VENTIL'       then RS:=ShortCut(65,[ssAlt]) else
  if Nom='VENTIL'       then RS:=ShortCut(65,[ssCtrl]) else
  if Nom='MODIFECHE'    then RS:=ShortCut(66,[ssCtrl]) else
  if Nom='LANCESAISIEBOR' then RS:=ShortCut(66,[ssAlt]) else
  if Nom='COMPLEMENT'   then RS:=ShortCut(67,[ssAlt]) else
  if Nom='INFOPOINTAGE' then RS:=ShortCut(67,[ssAlt]) else
  if Nom='COMBI'        then RS:=ShortCut(67,[ssAlt]) else
  if Nom='MONNAIE'      then RS:=ShortCut(68,[ssCtrl]) else
  if Nom='ECHE'         then RS:=ShortCut(69,[ssAlt]) else
  if Nom='CHERCHER'     then RS:=ShortCut(70,[ssCtrl]) else
  if Nom='GUIDE'        then RS:=ShortCut(71,[ssAlt]) else
  if Nom='CREERGUIDE'   then RS:=ShortCut(71,[ssCtrl]) else
  if Nom='CHANCEL'      then RS:=ShortCut(72,[ssAlt]) else
  if Nom='ZOOMDEVISE'   then RS:=ShortCut(73,[ssAlt]) else
  if Nom='ZOOMJOURNAL'  then RS:=ShortCut(74,[ssAlt]) else
  if Nom='PRORATA'      then RS:=ShortCut(75,[ssCtrl]) else
  if Nom='REF'          then RS:=ShortCut(76,[ssAlt]) else
  if Nom='LIBAUTO'      then RS:=ShortCut(76,[ssAlt]) else
  if Nom='LISTEPIECES'  then RS:=ShortCut(76,[ssCtrl]) else
  if Nom='COMPPIECE'    then RS:=ShortCut(76,[ssCtrl]) else
  if Nom='MODEPAIE'     then RS:=ShortCut(77,[ssAlt]) else
  if Nom='PIECEGC'      then RS:=ShortCut(77,[ssAlt]) else
  if Nom='MODIFGENERE'  then RS:=ShortCut(77,[ssAlt]) else
  if Nom='ZOOMIMMO'     then RS:=ShortCut(79,[ssCtrl]) else
  if Nom='PIECEHISTO'   then RS:=ShortCut(79,[ssCtrl]) else
  if Nom='IMPRIMER'     then RS:=ShortCut(80,[ssCtrl]) else
  if Nom='ETATPOINTAGE' then RS:=ShortCut(80,[ssCtrl]) else
  if Nom='DERNPIECES'   then RS:=ShortCut(80,[ssAlt]) else
  if Nom='CLEREPART'    then RS:=ShortCut(82,[ssCtrl]) else
  if Nom='ETATRAPPRO'   then RS:=ShortCut(82,[ssCtrl]) else
  if Nom='CHOIXREGIME'  then RS:=ShortCut(82,[ssCtrl]) else
  if Nom='REPART'       then RS:=ShortCut(82,[ssAlt]) else
  if Nom='CLEFBU'       then RS:=ShortCut(82,[ssAlt]) else
  if Nom='MODIFRIB'     then RS:=ShortCut(82,[ssAlt]) else
  if Nom='SOUSTOT'      then RS:=ShortCut(83,[ssCtrl]) else
  if Nom='SCENARIO'     then RS:=ShortCut(83,[ssAlt]) else
  if Nom='SOLDEGS'      then RS:=ShortCut(83,[ssAlt]) else
  if Nom='LANCESAISIE'  then RS:=ShortCut(83,[ssAlt]) else
  if Nom='ZOOMETABL'    then RS:=ShortCut(84,[ssAlt]) else
  if Nom='ZOOMETABL'    then RS:=ShortCut(84,[ssCtrl]) else
  if Nom='VENTILTYPE'   then RS:=ShortCut(86,[ssAlt]) else
  if Nom='MODIFTVA'     then RS:=ShortCut(86,[ssAlt]) else
  if Nom='VIDEPIECE'    then RS:=ShortCut(89,[ssCtrl]) else
  if Nom='ANNULERSEL'   then RS:=ShortCut(90,[ssCtrl]) else
  if Nom='APPLIQUER'    then RS:=ShortCut(VK_RETURN,[ssCtrl]) else
  if Nom='POINTE'       then RS:=ShortCut(VK_RETURN,[]) else
  if Nom='ABANDON'      then RS:=ShortCut(VK_ESCAPE,[]) else
  if Nom='CABANDON'     then RS:=ShortCut(VK_ESCAPE,[]) else
  if Nom='INSERT'       then RS:=ShortCut(VK_INSERT,[]) else
  if Nom='RAJOUTE'      then RS:=ShortCut(VK_INSERT,[ssShift]) else
  if Nom='DEL'          then RS:=ShortCut(VK_DELETE,[]) else
  if Nom='SDEL'         then RS:=ShortCut(VK_DELETE,[ssCtrl]) else
  if Nom='ENLEVE'       then RS:=ShortCut(VK_DELETE,[ssShift]) else
  if Nom='AIDE'         then RS:=ShortCut(VK_F1,[]) else
  if Nom='CAIDE'        then RS:=ShortCut(VK_F1,[]) else
  if Nom='PREV'         then RS:=ShortCut(VK_F3,[]) else
  if Nom='FIRST'        then RS:=ShortCut(VK_F3,[ssShift]) else
  if Nom='NEXT'         then RS:=ShortCut(VK_F4,[]) else
  if Nom='LAST'         then RS:=ShortCut(VK_F4,[ssShift]) else
  if Nom='ZOOM'         then RS:=ShortCut(VK_F5,[]) else
  if Nom='ZOOMECRITURE' then RS:=ShortCut(VK_F5,[]) else
  if Nom='ZOOMBUDGET'   then RS:=ShortCut(VK_F5,[]) else
  if Nom='ASSISTANT'    then RS:=ShortCut(VK_F5,[]) else
  if Nom='ZOOMABO'      then RS:=ShortCut(VK_F5,[ssShift]) else
  if Nom='ZOOMPIECE'    then RS:=ShortCut(VK_F5,[ssShift]) else
  if Nom='ZOOMCPTE'     then RS:=ShortCut(VK_F5,[ssShift]) else
  if Nom='ZOOMSECTION'  then RS:=ShortCut(VK_F5,[ssAlt]) else
  if Nom='SOLDE'        then RS:=ShortCut(VK_F6,[]) else
  if Nom='ZOOMMVTEFF'   then RS:=ShortCut(VK_F6,[]) else
  if Nom='SWAPPIVOT'    then RS:=ShortCut(VK_F8,[ssAlt]) else
  if Nom='SWAPEURO'     then RS:=ShortCut(VK_F8,[ssShift]) else
  if Nom='LIG'          then RS:=ShortCut(VK_F7,[]) else
  if Nom='CONTROLETVA'  then RS:=ShortCut(VK_F8,[]) else
  if Nom='GENERETVA'    then RS:=ShortCut(VK_F9,[]) else
  if Nom='VALIDE'       then RS:=ShortCut(VK_F10,[]) else
  if Nom='VALIDER'      then RS:=ShortCut(VK_F10,[]) else
  if Nom='CVALIDE'      then RS:=ShortCut(VK_F10,[]) else
  if Nom='VALIDESELECT' then RS:=ShortCut(VK_F10,[]) else
  if Nom='FERME'        then RS:=ShortCut(VK_ESCAPE,[]) else
  if Nom='LETTRAGE'     then RS:=ShortCut(Word('L'),[ssCtrl]) else
   ;
  if RS <> 0 then Result:=ShortCutToText(RS);
end;

{---------------------------------------------------------------------------------------}
procedure InitPopup(F : TForm);
{---------------------------------------------------------------------------------------}
var
  i,LeTag,k : integer ;
  T,N       : TMenuItem ;
  B,BT      : TBitBtn ;
  PP        : TPopupMenu ;
begin
  PP := TPopupMenu(F.FindComponent('POPS'));
  if PP = nil then Exit;

  {Purge des précédents items}
  PurgePopup(PP);
  
  {Création menu pop premier niveau}
  for i := 0 to F.ComponentCount - 1 do
    if ((F.Components[i].Tag > 0) and
        (TControl(F.Components[i]).Parent.Enabled) and
        (TControl(F.Components[i]).Parent.Visible)) then
      if IsButtonPop(F.Components[i]) then begin
        B:=TBitBtn(F.Components[i]) ;
        if ((B.Enabled) and ((B.Visible) or (B.Tag >= 1000))) then begin
          T := CreerLigPop(B, PP, True, 'P');
          PP.Items.Add(T);
        end ;
      end;

  {Création 2ème niveau si bouton "menu zoom"}
  for i:=0 to F.ComponentCount - 1 do
    if ((F.Components[i].Tag < 0) and (IsButtonPop(F.Components[i]))) then begin
      B := TBitBtn(F.Components[i]);
      if not B.Enabled then Continue;

      LeTag := -F.Components[i].Tag;
      T := CreerLigPop(B, PP, False, 'P');
      T.OnClick := nil;
      PP.Items.Add(T);

      for k:=0 to F.ComponentCount-1 do
        if ((F.Components[k].Tag = LeTag) and (IsButtonPop(F.Components[k]))) then begin
          BT := TBitBtn(F.Components[k]) ;
          if BT.Enabled then begin
            N := CreerLigPop(BT, T, True, 'P');
            T.Add(N);
          end;
        end;
    end;
  {$IFNDEF GIGI}
  AddMenuPop(PP, '', '');
  {$ENDIF}
end;

{$ENDIF EAGLSERVER}

{---------------------------------------------------------------------------------------}
function TAToStr(TA : TActionFiche) : string;
{---------------------------------------------------------------------------------------}
begin
         if (TA = taCreat)        then Result := 'TACREATE'
    else if (TA = taCreatEnSerie) then Result := 'TACREATEENSERIE'
    else if (TA = taCreatOne)     then Result := 'TACREATONE'
    else if (TA = taModif)        then Result := 'TAMODIF'
    else if (TA = taConsult)      then Result := 'TACONSULT'
    else if (TA = taModifEnSerie) then Result := 'TAMODIFENSERIE'
    else if (TA = taDuplique)     then Result := 'TADUPLIQUE'
                                  else Result := 'TACONSULT';
end;

{---------------------------------------------------------------------------------------}
function StrToTA(psz : string) : TActionFiche;
{---------------------------------------------------------------------------------------}
begin
         if (psz = 'TACREATE') or (psz='ACTION=CREATION')               then Result := taCreat
    else if (psz = 'TACREATEENSERIE') or (psz='ACTION=CREATIONENSERIE') then Result := taCreatEnSerie
    else if (psz = 'TACREATONE')                                        then Result := taCreatOne
    else if (psz = 'TAMODIF') or (psz='ACTION=MODIFICATION')            then Result := taModif
    else if (psz = 'TACONSULT') or (psz='ACTION=CONSULTATION')          then Result := taConsult
    else if (psz = 'TAMODIFENSERIE')                                    then Result := taModifEnSerie
    else if (psz = 'TADUPLIQUE')                                        then Result := taDuplique
                                                                        else Result := taConsult;
end;


{Récupère une action à partir d'une chaine
 Récupère une chaine à partir d'un mode d'action
{---------------------------------------------------------------------------------------}
procedure AvertirMultiTable(FTypeTable : string);
{---------------------------------------------------------------------------------------}
begin
  FTypeTable := Trim(UpperCase(FTypeTable));

  if FTypeTable = 'TTJOURNAL' then begin
    AvertirTable('ttJournal');
    AvertirTable('ttJalSaisie');
    AvertirTable('ttJalSansEcart');
    AvertirTable('ttJalBanque');
    AvertirTable('ttJalNonANouveau');
    AvertirTable('ttJalOD');
    AvertirTable('ttJalAnalytique');
    AvertirTable('ttJalRegul');
    AvertirTable('ttJalRegulDevise');
    AvertirTable('ttJalEcart');
    AvertirTable('ttJalNonBanque');
    AvertirTable('ttJalANouveau');
    AvertirTable('ttJalTVA');
    AvertirTable('ttJalGuide');
    AvertirTable('ttJalAnalAN');
    AvertirTable('ttJalVENOD');
    AvertirTable('ttBudJalSansCAT');
    AvertirTable('ttJALEFFET');
    AvertirTable('ttJALFOLIO');
    AvertirTable('ttJOURNAUX');
    AvertirTable('ttJALSCENARIO');
    AvertirTable('CPJOURNALIFRS');
  end

  else if FTypeTable = 'TTCATJALBUD' then begin
    AvertirTable('ttCatJalBud');
    AvertirTable('ttCatJalBud1');
    AvertirTable('ttCatJalBud2');
    AvertirTable('ttCatJalBud3');
    AvertirTable('ttCatJalBud4');
    AvertirTable('ttCatJalBud5');
  end

  {GCO - 20/09/2007 - FQ 21474 + Maj de toutes les tablettes sur EXERCICE}
  else if FTypeTable = 'TTEXERCICE' then begin
    AvertirTable('TTEXERCICE');
    AvertirTable('TTEXERCICEAPURGER');
    AvertirTable('TTEXERCICEBUDGET');
    AvertirTable('TTEXERCICEOUV');
    AvertirTable('TTEXERCICEOUVCPR');
    AvertirTable('TTEXERCICETRIE');
  end
  else
    AvertirTable(FTypeTable);
end;

{---------------------------------------------------------------------------------------}
function PrintSoldeFormate(TD, TC : Double; Decim : Integer; Symbole : string; AffSymb : Boolean; FormMont : string) : string ;
{---------------------------------------------------------------------------------------}
var
  LaValeur   : Double;
  LeResultat,
  StSym      : string;
  NegToPos   : Boolean;
begin
  LaValeur := 0;
  NegToPos := False;

  if (TD=TC) then begin
    LaValeur := TD - TC;
    LeResultat := FloatToStrF(LaValeur, ffNumber, 20, Decim);
    if AffSymb then PrintSoldeFormate := LeResultat + ' '+ Symbole + ' ' + ' '
               else PrintSoldeFormate := LeResultat + '  ' + ' '
  end
  else if Abs(TD) >= Abs(TC) then begin
    LaValeur := TD - TC;
    if (LaValeur < 0) and (FormMont = 'PC') then begin
      LaValeur := LaValeur * (-1);
      NegToPos := True;
    end;

    LeResultat := FloatToStrF(LaValeur, ffNumber, 20, Decim);
    if AffSymb then StSym := ' ' + Symbole
               else StSym := ' ';

    PrintSoldeFormate := LeResultat + StSym + ' ' + 'D';

    { Crédit positif }
    if FormMont = 'PC' then begin
      if NegToPos then PrintSoldeFormate := ' ' + LeResultat + StSym + ' ' + ' '
                  else PrintSoldeFormate := '-' + LeResultat + StSym + ' ' + ' ';
    end
    { Débit positif }
    else if FormMont='PD' then
      PrintSoldeFormate := LeResultat + StSym + ' ' + ' '
    { Debit parenthèse }
    else if FormMont='DP' then
      PrintSoldeFormate := '(' + LeResultat + ')' + StSym + ' ' + ' '
    { Crédit parenthèse }
    else if FormMont='CP' then
      PrintSoldeFormate := LeResultat + StSym + ' ' + ' ';

  end
  else if Abs(TD) < Abs(TC) then begin
    LaValeur := TC - TD;
    if (LaValeur < 0) and (FormMont = 'PD') then begin
      LaValeur := LaValeur * (-1);
      NegToPos := True;
    end;

    LeResultat := FloatToStrF(LaValeur, ffNumber, 20, Decim);

    if AffSymb then StSym := ' ' + Symbole
               else StSym := ' ';

    PrintSoldeFormate := LeResultat + StSym + ' ' + 'C';
    { Débit positif }
    if FormMont='PD' then begin
      if NegToPos then PrintSoldeFormate := ' ' + LeResultat + StSym + ' ' + ' '
                  else PrintSoldeFormate := '-' + LeResultat + StSym + ' ' + ' ';
    end
    { Crédit parenthèse }
    else if FormMont = 'CP' then
      PrintSoldeFormate := '(' + LeResultat + ')' + StSym + ' ' + ' '
    { Crédit positif }
    else if FormMont = 'PC' then
      PrintSoldeFormate := LeResultat + StSym + ' ' + ' '
    { Debit parenthèse }
    else if FormMont = 'DP' then
      PrintSoldeFormate := LeResultat + StSym + ' ' + ' ';
  end;

  if LaValeur = 0 then PrintSoldeFormate := '';
end;

{---------------------------------------------------------------------------------------}
function PrintSolde(TD, TC : Double; Decim : Integer; Symbole : string; AffSymb : Boolean) : string;
{---------------------------------------------------------------------------------------}

    {----------------------------------------------------------------------------}
    procedure RetoucheSolde(AffSymb : Boolean; DC : Char; Symbole : string; var LeResultat : string);
    {----------------------------------------------------------------------------}
    begin
      if AffSymb then LeResultat := LeResultat + ' ' + Symbole + ' ' + DC
                 else LeResultat := LeResultat + ' ' + DC;
    end;

var
  LaValeur   : Double;
  LeResultat : string;
  OnInverse  : Boolean;
begin
  LaValeur  := 0;
  OnInverse := False;

  if (TD = TC) then begin
    LaValeur   := TD - TC;
    LeResultat := FloatToStrF(LaValeur, ffNumber, 20, Decim);
    if AffSymb then PrintSolde := LeResultat + ' ' + Symbole + ' ' + ' '
               else PrintSolde := LeResultat + '  ' + ' ';
  end

  else if (Abs(TD) >= Abs(TC)) then begin
    LaValeur := TD - TC;
    if LaValeur < 0 then OnInverse := True;
      LeResultat := FloatToStrF(Abs(LaValeur), ffNumber, 20, Decim);
    if OnInverse then RetoucheSolde(AffSymb, 'C', Symbole, LeResultat)
                 else RetoucheSolde(AffSymb, 'D', Symbole, LeResultat);
  end
  else if Abs(TD) < Abs(TC) then begin
    LaValeur := TC - TD;
    if LaValeur < 0 then OnInverse := True;
      LeResultat := FloatToStrF(Abs(LaValeur), ffNumber, 20, Decim);
    if OnInverse then RetoucheSolde(AffSymb, 'D', Symbole, LeResultat)
                 else RetoucheSolde(AffSymb, 'C', Symbole, LeResultat);
  end;

  if LaValeur = 0 then PrintSolde := ''
                  else PrintSolde := LeResultat;
end;

{---------------------------------------------------------------------------------------}
function PrintEcart(TD, TC : Double; Decim : Integer; DebitPos : Boolean) : string;
{---------------------------------------------------------------------------------------}
var
  LaValeur : Double;
begin
  LaValeur := 0;
  if (TD = TC) then begin
    LaValeur := Arrondi(TD - TC, Decim);
    Result := FloatToStrF(LaValeur, ffNumber, 20, Decim) + '   ';
  end
  else if Abs(TD) >= Abs(TC) then begin
    LaValeur := Arrondi(TD - TC, Decim);
    if not DebitPos then LaValeur := -1 * LaValeur;
    Result   := FloatToStrF(LaValeur, ffNumber, 20, Decim) + '   ';
  end
  else if Abs(TD) < Abs(TC) then begin
    LaValeur := Arrondi(TC - TD, Decim);
    if DebitPos then LaValeur := -1 * LaValeur;
    Result := FloatToStrF(LaValeur, ffNumber, 20, Decim) + '   ';
  end;

  if LaValeur = 0 then Result := '';
end;

{---------------------------------------------------------------------------------------}
function AfficheMontant(Formatage, LeSymbole : string; LeMontant : Double; OkSymbole : Boolean) : string;
{---------------------------------------------------------------------------------------}
begin
  if OkSymbole then begin
    if LeMontant = 0 then AfficheMontant := ''
                     else AfficheMontant := FormatFloat(Formatage + ' ' + LeSymbole, LeMontant);
  end
  else begin
   if LeMontant = 0 then AfficheMontant := ''
                    else AfficheMontant := FormatFloat(Formatage, LeMontant);
   end;
end;

{---------------------------------------------------------------------------------------}
function TotDifferent(X1, X2 : Double) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result:= not (Abs(X1 - X2) < 0.000001);
end ;

{ Déformate les montants
{---------------------------------------------------------------------------------------}
function MoinsSymbole(LeSolde : string) : Double;
{---------------------------------------------------------------------------------------}
var
  i : Integer;
  Alors : string;
begin
  Alors := '';
  if LeSolde = '' then LeSolde := '0';
  for i:=1 to Length(LeSolde) do
    if LeSolde[i] in ['0'..'9', '.', ',', '-'] then Alors := Alors + LeSolde[i];

  MoinsSymbole := StrToFloat(Alors);
end;

{22/10/07 : FQ 21706 : transforme un fichier sans retour de chariot en fichier avec
{---------------------------------------------------------------------------------------}
function TransformeFichier(NomFichier : string; lgEnreg : Integer = 120) : string;
{---------------------------------------------------------------------------------------}
var
  LSource : TStringList;
  LDest   : TStringList;
  Enreg   : string;
  BakName : string;
  Rep     : string;
  lgFic   : Integer;
  Indice  : Integer;
begin
  Result  := '';
  LSource := TStringList.Create;
  LDest   := TStringList.Create;
  try
    LSource.LoadFromFile(NomFichier);
    if LSource.Count = 1 then begin
      Enreg := LSource[0];
      lgFic  := Length(Enreg);
      if (lgFic mod lgEnreg) = 0 then begin
        Indice := 1;
        while lgFic > 0 do begin
          LDest.Add(Copy(Enreg, Indice, lgEnreg));
          Indice := Indice + lgEnreg;
          lgFic := lgFic - lgEnreg;
        end;

        Rep := ExtractFilePath(NomFichier) + 'MULTI';
        SysUtils.ForceDirectories(Rep);
        BakName := ExtractFileName(NomFichier);
        BakName := ChangeFileExt(BakName, '.~REL');
        LSource.SaveToFile(Rep + '\' + BakName);

        LDest.SaveToFile(NomFichier);
      end
      else
        Result := TraduireMemoire('Le fichier n''est pas cohérent :') + NomFichier;
    end;
  finally
    FreeAndNil(LSource);
    FreeAndNil(LDest);
  end;
end;

{Decoupe un fichier contenant plusieurs "ensembles"
{---------------------------------------------------------------------------------------}
procedure DecoupeFichier(NomFichier, Deb, Fin, Ext : string);
{---------------------------------------------------------------------------------------}
var
  LSource : TStringList;
  LDest   : TStringList;
  n       : Integer;
  p       : Integer;
  k       : Integer;
  NomRel  : string;
  Name    : string;
  Rep     : string;
begin
  k := LastDelimiter('.\:', NomFichier);
  if (k = 0) or (NomFichier[k] <> '.') then k := MaxInt;
  Name := Copy(NomFichier, 1, k - 1);

  LSource  := TStringList.Create;
  LDest    := TStringList.Create;
  try
    LSource.LoadFromFile(NomFichier);
    p := 1;
    for n := 0 to LSource.Count - 1 do begin
      if (n > 0) and (Copy(LSource[n], 1, Length(Deb)) = Deb) then begin
        NomRel := Name + '_' + IntToStr(p) + '.' + Ext;
        LDest.SaveToFile(NomRel);
        LDest.Clear;
        Inc(p);
      end;
      LDest.Add(LSource[n]);
    end;

    if p > 1 then begin
      NomRel := Name + '_' + IntToStr(p) + '.' + Ext;
      LDest.SaveToFile(NomRel);

      if Ext = 'DAT' then begin
        Rep := ExtractFilePath(NomFichier) + 'MULTI';
        SysUtils.ForceDirectories(Rep);
        Name := ExtractFileName(NomFichier);
        Name := ChangeFileExt(Name, '.~MUL');
        LSource.SaveToFile(Rep + '\' + Name);
        SysUtils.DeleteFile(NomFichier);
      end;
    end;
  finally
    FreeAndNil(LSource);
    FreeAndNil(LDest);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
function CreateTobParamCWA( vStNameAppli : HString ) : Tob;
begin
  Result := Tob.Create('$PARAM', nil, -1);
  Result.AddChampSupValeur('USERLOGIN' ,  V_PGI.UserLogin ) ;
  //Result.AddChampSupValeur('INIFILE' ,    HalSocIni ) ;
  Result.AddChampSupValeur('PASSWORD' ,   V_PGI.Password ) ;
  Result.AddChampSupValeur('DOMAINNAME',  V_PGI.DomainName ) ;
  Result.AddChampSupValeur('DATEENTREE',  V_PGI.DateEntree ) ;
  Result.AddChampSupValeur('DOSSIER'   ,  V_PGI.NoDossier ) ;
  Result.AddChampSupValeur('BASECOMMUNE', V_PGI.DBName = V_pgi.DefaultSectionDBName);
  Result.AddChampSupValeur('APPLICATION', vStNameAppli);
  Result.AddChampSupValeur('ERROR', '');
end;

////////////////////////////////////////////////////////////////////////////////

end.

