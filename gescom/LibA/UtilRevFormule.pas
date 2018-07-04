unit UtilRevFormule;

interface

  uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
  Forms, Dialogs, StdCtrls, ExtCtrls,grids,Buttons,printers, Utob,
  hctrls, hent1,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  UAFO_RevPrixCalculCoef;

{  TFORMULE = class
  private
    fStOperateur  : String;
    fStFormule    : String;
    fStElemPrec   : String;
    fStElemSuiv   : String:
    fInNiveau     : Integer;
  end;
}

  const max=20;

  type
    str2  = string[2];
    str3  = string[3];
    str4  =string[4];
    str5  = string[5];
    str6  = string[6];
    str7  = string[7];
    str8  = string[8];
    str9  = string[9];
    str10 =string[10];
    str12 = string[12];
    str15 = string[15];
    str20 = string[20];
    str30 = string[30];
    str32 = string[32];
    str50 = string[50];
    str80 = string[80];
    str132 = string[132];

  type
    schema = array[1..3] of string;

  function b(Num : integer) : String;
  function cop(const s: string;Len:byte):string;
  function rien(const s: string): boolean;
  function cope(const s:string;i:byte):string;
  function majus(const S : String): String;
  function sup_blanc(const s: string;type_sup:byte):string;
  {comparaison d'une valeur entre deux autres}
  function test_int(i,i1,i2:longint): boolean ;
  {comparaison en majuscule à l'identique de 2 chaînes}
  function str_egal(s1,s2:string):boolean;
  function test_str(s,s1,s2: string): boolean ;
  function ces_num(var s:string):boolean;
  function fract(vl:double;deci:byte):double;
  function valli(const s :str132) :longint;
  function vallv(const s :str132) :double;
  function in_st(Num, Len : longint) : String;
  { Changes integer en string }
  function re_st(Num : double; Len, nbrdec : Word) : String;
  { Changes a real to a string }
  function blanc_zero(const s:str80):str80;
  function num_str(i,l:longint):str50;
  function inz_st(Num, Len : longint): String;
  function binz_st(Num, Len : integer): String;
  function size_form(f: Tform;var CoeffScreenVert,CoeffScreenHoriz : Real):boolean;
  procedure ImprimeEcran(form: TForm);
  FUNCTION lect_nbr(var s:string;var z:smallint):double;
  function lect_str(var s:string;var z:smallint;car:char):string;

  procedure TraitementFormule(expression : string; var pMemoDetail, pMemoEdition : TMemo);
//  function FormuleEdition(pInNumLigne : Integer; pStAffaire, pStForCode : String; var pSLFormule, pSLDetail : TStringList) : Boolean;
  function FormuleEdition(pInNumLigne : Integer; pStAffaire, pStForCode : String) : String;
  function FormuleDetail(pInNumLigne : Integer; pStAffaire, pStForCode : String) : String;
  procedure RemplaceIndice(pStForCode, pStAffaire : String; var pStMemo : String);
  function niveau2(pStFormule : String; pMemo : TMemo; var pStNewFormule : String) : Integer;
  procedure MemoDetailEdition(pMemoDetail : TMemo; var pSlDetailEdition : TStringList);

implementation

function b(Num : integer): String;
var
  s : String; { chaîne à remplir de blancs}
begin
  if num<=0 then begin result:='';exit;end;
  SetLength(S,Num);
  FillChar(S[1], Num, ' ');
  result := S;
end;

function re_st(Num : double ; Len, nbrdec : Word): String;
var     {réel=>string}
  s : String;
begin
  Str(Num: Len:nbrdec, S);
  if (num=0)  then
    s := b(len);
  result := s;
end;

function rien(const s: string): boolean;
var        {la chaine est t'elle vide ou blanche}
  i: word; { indice de boucle sur la la chaîne}
begin
  result := true;
  if length(s) > 0 then
    for i :=1 to length(s) do
      if s[i] <> ' ' then
        result := false;
end;

{------------------------------------------------------------------------------
  Procédure de copie des x premiers caractères d'une chaîne
  ------------------------------------------------------------------------------}
function cop(const s: string;Len:byte):string;
var
  i : word; { indice de boucle sur la chaîne}
  chaine : string;
begin
  chaine := s;
  if length(s) = 0 then
    result := b(len)
  else begin
    for i:=1 to length(s) do
      if chaine[i] = #0 then
        chaine[i] := ' ';
    if length(chaine) <= len then
      chaine := chaine+b(len-length(chaine));
    result := copy(chaine,1,len);
    end;
end;
{------------------------------------------------------------------------------
  Procédure de copie des x derniers caractères d'une chaîne
  ------------------------------------------------------------------------------}
function cope(const s: string;i:byte):string;
begin
  if (length(s) > 0) and (i < length(s)) then
    result := copy(s,length(s)-i+1,i)
  else
    result := s;
end;
{------------------------------------------------------------------------------
  Procédure de conversion en majuscule d'une chaîne dos
  ------------------------------------------------------------------------------}
function majus(const S : String): String;
var
  Counter : Word; { indice de boucle sur la longueur de la chaîne}
begin
  result := s;
  if length(s) > 0 then
    for Counter := 1 to Length(S) do begin
      result[Counter] := Upcase(S[Counter]);
      case result[counter] of
        'é','è','ê',#200..#203: result[counter]:='E';
        'à','â',#192..#198    : result[counter]:='A';
        'û','ù',#218..#220    : result[counter]:='U';
        'ç'                   : result[counter]:='C';
        'ï','î',#204..#207    : result[counter]:='I';
        'ô','ö',#210..#214    : result[counter]:='O';
        end;
      end;
end;

 function str_egal(s1,s2: string):boolean;
begin
 {comparaison de 2 chaines qqsoit la casse}
  str_egal := majus(sup_blanc(s1,1)) = majus(sup_blanc(s2,1));
end;

{suppression des blancs, devant (0) derrière(1) ou tous (2)}
function sup_blanc(const s: string;type_sup:byte):string;
var
  chaine : string;
begin
  chaine := s;
  case type_sup of
    0 :
      while (Length(chaine) > 0) and (chaine[1] <= ' ') do
        Delete(chaine, 1, 1);
    1 :
      while (Length(chaine) > 0) and (chaine[Length(chaine)] <= ' ') do
        delete(chaine,length(chaine),1);
    2 : begin
          while pos(' ',chaine) <> 0 do
            delete(chaine,pos(' ',chaine),1);
          while pos(chr(0),chaine) <> 0 do
            delete(chaine,pos(chr(0),chaine),1);
          end;
    3 : begin
          while (Length(chaine) > 0) and (chaine[1] <= ' ') do
            Delete(chaine, 1, 1);
          while (Length(chaine) > 0) and (chaine[Length(chaine)] <= ' ') do
            delete(chaine,length(chaine),1);
          end;
    end;
  result := chaine;
end;

function test_str(s,s1,s2: string): boolean ;
{chaine comprise entre deux autres}
begin
  result := (sup_blanc(s,1) >= sup_blanc(s1,1)) and
    (sup_blanc(s,1) <=sup_blanc(s2,1) );
end;




{------------------------------------------------------------------------------
  Procédure permettant de contrôler si une chaîne est numérique
  ------------------------------------------------------------------------------}
function ces_num(var s: string):boolean;
var
  j : byte; { indice de boucle sur la chaîne}
  i : byte; { présence d'un point décimal}
  k : byte; { présence du signe +}
  l : byte; { présence du signe -}
  m : byte; { position du point décimal}
begin
  result := true; {false; Lek 050898}
  { contrôle des caractères}
  if length(s) > 0 then begin
    {result := true; Lek 050898}
    for j:=1 to length(s) do
      if not (s[j] in ['0'..'9','+','-',' ','.']) then
        result := false;

    { mise en forme de la chaîne}
    if result then begin
      i:=0;
      k:=0;
      l:=0;
      m:=0;
      for j:=1 to length(s) do
        case s[j] of
          '.' : begin
                  m:=j;
                  inc(i);
                  end;
          '+' : inc(k);
          '-' : inc(l);
          end;
      if (i=1) and (s[m+1]=' ') then
        s[m+1] := '0';
      if (i>1) or (k>1) or (l>1) or (k+l>1) then
        result := false;
      end;
    end;
end;




function fract(vl:double;deci:byte):double;
var pm,j,i:integer;k:double;
{calcul de l'arrondi à deci décimales}
begin
if vl=0 then begin fract:=0;exit;end;
i:=deci;if i>4 then i:=4;if i<0 then i:=2;j:=1;
case i of
  0:begin k:=5E-1;j:=1;end;
  1:begin k:=5E-2;j:=10;end;
  2:begin k:=5E-3;j:=100;end;
  3:begin k:=5E-4;j:=1000;end;
  4:begin k:=5E-5;j:=10000;end;
  else k:=0;
  end;
pm:=1;if vl<0 then begin vl:=-vl;pm:=-1;end;
result:=pm*((vl+k)*j-frac((vl+k)*j))/j;
end;

{valeur entière d'une chaine}
function valli(const s :str132) :longint;
var i:longint;j:integer;
    ss:string;
begin
  ss:=sup_blanc(s,2);
  valli:=0; if ss ='' then exit;
  if not ces_num(ss) then begin
    exit;
    end;
  if ss[length(ss)] in ['+','-'] then ss:=ss[length(ss)]+cop(ss,length(ss)-1);
  val(ss,i,j);
  if j<>0 then valli:=0
          else valli:=i;
end;

{valeur réelle d'une chaine}
function vallv(const s :str132) :double;
var vl:extended;j:integer;
    ss:^string;
begin
  new(ss);
  ss^:=sup_blanc(s,2);
  vallv:=0;if rien(s) then exit;
  if not ces_num(ss^) then begin
    dispose(ss);
    exit;
    end;
  if ss^[length(ss^)] in ['+','-'] then ss^:=ss^[length(ss^)]+cop(ss^,length(ss^)-1);
  val(ss^,vl,j);
  if j<>0 then vallv:=0
          else vallv:=vl;
  dispose(ss);
end;


function bool_o_n(x:boolean) : char;
begin if x then bool_o_n:='O' else bool_o_n:='N';end;


{------------------------------------------------------------------------------
  Procédure de conversion d'un numérique entier en chaîne
  ------------------------------------------------------------------------------}
function in_st(Num, Len : longint): String;
var
  S : String; { chaîne de travail}
begin
  if len <= 0 then
    len := 1;
  Str(Num: Len, S);
  if num = 0 then
    s := b(len-1)+'0';
  result := S;
end;

{------------------------------------------------------------------------------
  Procédure permettant de remplacer les blancs par des 0 dans une chaîne
  ------------------------------------------------------------------------------}
function blanc_zero(const s: str80):str80;
var
  ss : string; { chaîne de travail}
begin
  ss := s;
  while pos(' ',ss) <> 0 do
    ss[pos(' ',ss)] := '0';
  result := ss;
end;

{------------------------------------------------------------------------------
  Procédure de conversion d'un numérique entier en chaîne en rendant 0 si rien
  ------------------------------------------------------------------------------}
function inz_st(Num, Len : integer): String;
var
  S : String;
begin
  if len <= 0 then
    len := 1;
  Str(Num: Len, S);
  if num = 0 then begin
    s := b(len);
    s[len] := '0';
    end;
  result := S;
end;

{------------------------------------------------------------------------------
  Procédure de conversion d'un numérique entier en chaîne
  en rendant 0 si rien mais sans blanc
  ------------------------------------------------------------------------------}
function binz_st(Num, Len : integer): String;
var
  S : String;
begin
  if len <= 0 then
    len := 1;
  Str(Num: Len, S);
  if num = 0 then begin
    s := b(len);
    s[len] := '0';
    end;
  result := sup_blanc(s,0);
end;

{------------------------------------------------------------------------------
  Procédure de conversion d'un numérique entier en chaîne avec conversion des
  blancs en 0
  ------------------------------------------------------------------------------} 
function num_str(i,l: integer):str50;
begin
  result := blanc_zero(in_st(i,l));
end;

 {entier entre deux autres ?}
function test_int(i,i1,i2:integer): boolean ;
begin
test_int:=(i >= i1) and (i<=i2);
end;

function sm(v:double):string;
begin
result:=sup_blanc(re_st(v,14,2),2);if result='' then result:='0.0';
end;


function size_form(f: Tform;var CoeffScreenVert,CoeffScreenHoriz : Real):boolean;
var
  NumCompon: Integer;
  CoeffMinus: Real;
const
  ComplBarTach=0.9;
begin
  {permet l'agrandissement automatique de la fenetre en fonction de l'écran
    à mettre dans init de la fiche avec déclaration de 2 var dans l'objet}
  with f do begin
    try
      CoeffScreenHoriz:= (Screen.Width / (Width+1));
      CoeffScreenVert := (Screen.Height* ComplBarTach / (Height+1));
      ClientWidth:= Round(ClientWidth * CoeffScreenHoriz);
      ClientHeight:= Round(ClientHeight * CoeffScreenVert);
      if (CoeffScreenHoriz<=CoeffScreenVert) then
        CoeffMinus:=CoeffScreenHoriz
      else
        CoeffMinus:=CoeffScreenVert;

      for NumCompon := 0 To (ComponentCount-1) do
        if (Components[NumCompon] Is TWinControl) or
          (Components[NumCompon] Is TSpeedButton) then
          if (Components[NumCompon] Is TWinControl) then
            with TWincontrol(Components[NumCompon]) do begin
              Left:= Round(Left*CoeffScreenHoriz);
              Top:= Round(Top*CoeffScreenVert);
              Width:= Round(Width*CoeffScreenHoriz);
              Height:= Round(Height*CoeffScreenVert);
              end
            else
              with TSpeedButton(Components[NumCompon]) do begin
                Left:= Round(Left*CoeffScreenHoriz);
                Top:= Round(Top*CoeffScreenVert);
                Width:= Round(Width*CoeffScreenHoriz);
                Height:= Round(Height*CoeffScreenVert);
                end;

      for NumCompon := 0 To (ComponentCount-1) do begin
        if (Components[NumCompon] is TEdit) then begin
          TEdit(Components[NumCompon]).AutoSize:=False;
          TEdit(Components[NumCompon]).Font.Size:= Round(CoeffMinus*TEdit(Components[NumCompon]).Font.Size);
          end;
        if (Components[NumCompon] is TStringGrid) then begin
          {TStringGrid1(Components[NumCompon]).Font.Size:= Round(CoeffMinus*TStringGrid1(Components[NumCompon]).Font.Size)}
          ;
          TStringGrid(Components[NumCompon]).DefaultRowHeight:=
            Round(CoeffScreenVert*TStringGrid(Components[NumCompon]).DefaultRowHeight);
          {for NumCol:=0 to TStringGrid1(Components[NumCompon]).ColCount-1 do
            TStringGrid1(Components[NumCompon]).ColWidths[NumCol]:=
            Round(CoeffScreenHoriz*TStringGrid1(Components[NumCompon]).ColWidths[NumCol])}
          end;
        end;
      result:=true;
    except
      CoeffScreenHoriz := 1;
      CoeffScreenVert := 1;
      result:=false;
      end;
    end;
end;
   {impression de la fenetre supérieure}
  { A l'appel, form doit contenir la fenêtre à imprimer}
  procedure ImprimeEcran(form: TForm);
  var
    DC            : HDC; { Contexte de périphérique}
    MemDC         : HDC; { Contexte de périphérique compatible à DC}
    MemBitmap     : HBitmap; { Handle de bitmap mémoire}
    OldMemBitmap  : HBitmap; { Ancien Handle de bitmap sélectionné}
    IsDCPalDevice : BOOL; { Indicateur de correspondance de palette}
    PPal          : PLOGPALETTE; { Palette en cours}
    Pal           : HPalette; { Handle de palette}
    OldPal        : HPalette; { Ancien Handle de palette}
    HDibHeader    : Thandle; { Handle global de structure d'info bitmap}
    PDibHeader    : Pointer; { Pointeur d'accès à HDIBHeader}
    HBits         : Thandle; { Handle globale de la zone de réception}
    PBits         : Pointer; { Pointeur sur handle HBits}
    I             : Integer; { Compteur de boucle}
    ScaleX,ScaleY : double;
  begin
    // SECTION DE CREATION DU BITMAP DANS UN CONTEXTE DE PERIPHERIQUE INDEPendANT
    { Demande de contexte d'écran générique}
    DC:=GetDC(0);
    { Création d'un contexte de périphérique compatible à DC.}
    MemDC:=CreateCompatibleDC(DC);
    { Création d'un structure bitmap compatible avec le contexte de périphérique}
    { écran. Ce bitmap mémoire aura la largeur et la hauteur de la fenêtre que l'on}
    { veut imprimer.}
    MemBitmap:=CreateCompatibleBitmap(DC,form.width,Form.height);
    { Sélection du bitmap mémoire dans le contexte de périphérique compatible.}
    { L'ancien bitmap mémoire est sauvegardé dans OldMemBitmap.}
    OldMemBitmap:=SelectObject(MemDC,MemBitmap);
    { Préparation de la palette à utiliser le cas échéant}
    IsDCPalDevice:=False;
    PPal:=nil;
    Pal:=0;
    OldPal:=0;
    if (GetDeviceCaps(DC,RASTERCAPS) and RC_PALETTE) = RC_PALETTE then begin
      GetMem(PPal,SizeOf(TLOGPALETTE)+(255*SizeOf(TPALETTEENtry)));
      FillChar(PPal^,SizeOf(TLOGPALETTE)+(255*sizeof(TPALETTEENtry)),#0);
      PPal^.PalVersion:=$300;
      PPal^.palNumEntries:=GetSystemPaletteEntries(DC,0,256,PPal^.palPalEntry);
      if PPal^.PalNumEntries <> 0 then begin
        Pal:=CreatePalette(PPal^);
        OldPal:=SelectPalette(MemDC,Pal,False);
        IsDCPalDevice:=True;
        end
      else
        FreeMem(PPal,SizeOf(TLOGPALETTE)+(255*SizeOf(TPALETTEENtry)));
      end;
    { Copie du contexte écran vers le contexte compatible}
    BitBlt(MemDC,0,0,form.Width,Form.Height,DC,Form.Left,Form.Top,SrcCopy);
    { Restauration de la palette}
    if IsDCPalDevice then begin
      SelectPalette(MemDC,OldPal,False);
      DeleteObject(Pal);
      end;
    { restauration du bitmap mémoire}
    SelectObject(MemDC,OldMemBitmap);
    { Libération du contexte mémoire}
    DeleteDC(MemDC);
    { Allocation mémoire pour la structure DIB}
    HDibHeader:=GlobalAlloc(GHND,SizeOf(TBITMAPINFO)+(SizeOf(TRGBQUAD)*256));
    { Demande de pointeur sur structure}
    PDibHeader:=GlobalLock(HDibHeader);
    { Remplissage de la structure DIP avec les infos}
    FillChar(PDibHeader^,SizeOf(TBITMAPINFO)+(SizeOf(TRGBQUAD)*256),#0);
    PBITMAPINFOHEADER(PDibHeader)^.BiSize:=SizeOf(TBITMAPINFOHEADER);
    PBITMAPINFOHEADER(PDibHeader)^.BiPlanes:=1;
    PBITMAPINFOHEADER(PDibHeader)^.BiBitCount:=8;
    PBITMAPINFOHEADER(PDibHeader)^.BiWidth:=form.Width;
    PBITMAPINFOHEADER(PDibHeader)^.BiHeight:=form.Height;
    PBITMAPINFOHEADER(PDibHeader)^.BiCompression:=BI_RGB;
    { Remplissage du bitmap pointé par PDibHeader}
    GetDIBits(DC,MemBitmap,0,form.Height,nil,TBitmapInfo(PDibHeader^),DIB_RGB_COLORS);
    { Allocation mémoire des octets}
    HBits:=GlobalAlloc(GHND,PBitmapInfoHeader(PDibHeader)^.BiSizeImage);
    { Demande de pointeur sur structure}
    PBits:=GlobalLock(HBits);
    { On appelle de nouveau la fonction mais cette fois pour les octets!}
    GetDIBits(DC,MemBitmap,0,form.Height,PBits,PBitmapInfo(PDibHeader)^,DIB_RGB_COLORS);
    { Transfert de palette}
    if IsDCPalDevice then begin
      {$R-}
      { La directive est obligatoire car les structures BitMapInfo sont généralement}
      { représentées par des array[0..0] of ...}
      for I:=0 to PPal^.PalNumEntries-1 do begin
        PBitmapInfo(PDibHeader)^.bmiColors[I].rgbRed:=PPal^.palPalEntry[I].peRed;
        PBitmapInfo(PDibHeader)^.bmiColors[I].rgbGreen:=PPal^.palPalEntry[I].peGreen;
        PBitmapInfo(PDibHeader)^.bmiColors[I].rgbBlue:=PPal^.palPalEntry[I].peBlue;
        end;
      {$R+}
      FreeMem(PPal,SizeOf(TLOGPALETTE)+(255*SizeOf(TPALETTEENtry)));
      end;
    { Libération du contexte d'écran}
    ReleaseDC(0,DC);
    { Supression du bitmap}
    DeleteObject(MemBitmap);

    // SECTION D'IMPRESSION DU BITMAP INDEPendANT
    { Début d'impression}
    Printer.begindoc;
    { formattage de la zone d'impression}
    if Printer.PageWidth < Printer.PageHeight then begin
      ScaleX:=Printer.PageWidth;
      ScaleY:=form.Height*(Printer.PageWidth/Form.Width);
      end
    else begin
      ScaleX:=form.Width*(Printer.PageHeight/Form.Height);
      ScaleY:=Printer.PageHeight;
      end;
    { Préparation de la palette à utiliser le cas échéant}
    IsDCPalDevice:=False;
    if (GetDeviceCaps(Printer.Canvas.Handle,RASTERCAPS) and RC_PALETTE)
      = RC_PALETTE then begin
      { Création de la palette à partir de la structure DIB}
      GetMem(PPal,SizeOf(TLOGPALETTE)+(255*SizeOf(TPALETTEENtry)));
      FillChar(PPal^,SizeOf(TLOGPALETTE)+(255*SizeOf(TPALETTEENtry)),#0);
      PPal^.palVersion:=$300;
      PPal^.palNumEntries:=256;
      { Transfert}
      for i:=0 to PPal^.PalNumEntries-1 do begin
        PPal^.palPalEntry[I].peRed:=PBitmapInfo(PDibHeader)^.bmiColors[I].rgbRed;
        PPal^.palPalEntry[I].peGreen:=PBitmapInfo(PDibHeader)^.bmiColors[I].rgbGreen;
        PPal^.palPalEntry[I].peBlue:=PBitmapInfo(PDibHeader)^.bmiColors[I].rgbBlue;
        end;
      Pal:=CreatePalette(PPal^);
      FreeMem(PPal,SizeOf(TLOGPALETTE)+(255*SizeOf(TPALETTEENtry)));
      OldPal:=SelectPalette(Printer.Canvas.Handle,Pal,False);
      IsDCPalDevice:=True;
      end;
    { Envoi des octets au contexte de périphérique d'impression}
    StretchDiBits(Printer.Canvas.Handle,0,0,Round(ScaleX),Round(ScaleY),0,0
      ,form.Width,Form.Height,PBits,PBitmapInfo(PDibHeader)^,DIB_RGB_COLORS,SRCCOPY);
    { Restauration de la palette}
    if IsDCPalDevice then begin
      SelectPalette(Printer.Canvas.Handle,OldPal,False);
      DeleteObject(Pal);
      end;
    { Libération de la mémoire alouée}
    GlobalUnlock(HBits);
    GlobalFree(HBits);
    GlobalUnlock(HDibHeader);
    GlobalFree(HDibHeader);
    { Fin d'impression}
    Printer.enddoc;
  end;


FUNCTION lect_nbr(var s:string;var z:smallint):double;
var j:smallint;
    sous_chaine:string;
begin j:=z;
  while (j<length(s)+1) and (s[j] in ['0'..'9','.'])  do j:=j+1;
  sous_chaine:=copy(s,z,j-z);
  if pos('.',sous_chaine)<>0 then
    result:=vallv(sous_chaine)
  else
    result:=valli(sous_chaine);
  z:=j-1;
end;

function lect_str(var s:string;var z:smallint;car:char):string;
var j:integer;
    sous_chaine:string;
begin
  j:=z+1;{il faut sauter le car ouverture }
  while (j<length(s)) and (s[j] <>car)  do j:=j+1;
  sous_chaine:=copy(s,z+1,j-z-1);
  z:=j; {il faut sauter le car de fermeture }
  result:=sous_chaine;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 03/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TraitementFormule(expression : string; var pMemoDetail, pMemoEdition : TMemo);
var
  sous_chaine : array[1..max] of string;{parenthèses internes}
  val_chaine  : array[1..max] of double;
  i, NN       : smallint; {indice de mémorisation des sous chaines}
  s           : string;

  // function 1
  {on recherche les parenthèses les plus intèrieures pour les remplacer par
   des références à un tableau de sous chaînes, dès que l'on rencontre une
   parenthèse fermante on revient à la dernière parenthèse ouvrante}
  function suppression_parenthese_interne(s:string):string;
  var j,dern_ouv,prem_ferm : smallint;
  begin
    j:=0;dern_ouv:=0;
    while j<length(s) do
    begin
      inc(j);
      case s[j] of
        '(':begin
            dern_ouv:=j+1;
            end;
        ')':begin
            prem_ferm:=j-1;
            sous_chaine[NN]:=copy(s,dern_ouv,prem_ferm-dern_ouv+1);
            result:=cop(s,dern_ouv-2)+'@'+num_str(NN,2)+cope(s,length(s)-prem_ferm-1);
            exit;
            end;
      end;
    end;
  end;
  {(1+34-(23/5+7)/((5+6)*(7-8)+9))/(23+5/6) }
  {
  0.15+(0.55*([SICMO2]/[SICMO2°]))+0.05*([PSDA3]/[PSDA3°])+0.125*([GOHL1870]/[GOHL1870°])+0.125*([VEHBMSC]/[VEHBMSC°])
  }

  // function 2
  {fonction simplifiée car à ce niveau il n'y a plus de parenthèses}
  function calcul_chaine(s:string):double;
  var ope:char;
      mt,total:double;
      j:smallint;

    // function 21
    function valeur_indice(ind:string):double;
    begin
    result:=1;
    end;

    // function 22
    procedure calcul(ope:char;mtc:double);
    begin
    case ope of
      '+',' ':begin total:=total+mt;mt:=mtc;end;
      '-':begin total:=total+mt;mt:=-mtc;end;
      '*':mt:=mt*mtc;
      '/':if mtc<>0 then mt:=mt/mtc;
      end;
    end;

    begin   {calcul_chaine}
    s:=s+' ';j:=0;total:=0;ope:=' ';mt:=0;
    while j<length(s) do begin
      inc(j);
      case s[j] of
        '0'..'9','.'   :calcul(ope,lect_nbr(s,j));
        '['            :calcul(ope,valeur_indice(lect_str(s,j,']')));
        '+','-','/','*':ope:=s[j];
        '@'            :begin inc(j);calcul(ope,val_chaine[trunc(lect_nbr(s,j))]);end;
        end;
      end;{while}
    result:=total+mt;
    end;

  //function 3
  function car(n:smallint):string;
  var i:smallint;
  begin
  result:=''; i:=0;
  while i<n do begin result:=result+'-';inc(i);end;
  end;

  // function 4
  function inv(n:smallint):smallint;
  begin
  if n=1 then result:=3 else result:=1;
  end;

  // function 5
  {cette fonction lit une chaine dite simple c'est à dire sans parenthèse
  en effet s'il y une parenthèse elle a été remplacé par l'appel à une
  sous chaine @nn en mode récursif
  on cherche à remplir dans cette fonction les trois lignes à éditer
  pour une formule
    ligne 1 contient les numérateurs
    "     2 contient les traits de division
    "     3 contient le dénominateurs
  si on appelle une sous chaine par @nn on recoit en retour 3 morceaux
  de ligne (dans l_s) que l'on peut concaténer à l_c en cours
  }
  function traitement_edition(ss:string):schema;
  var l_c,l_s:schema;
      niv, {position sur ligne 1, 3  }
      i,
      avant,{position avant lecture élément}
      l_element {longueur du dernier élément}:smallint;
     {l1    1+2+5*(5 +9)-12*9
      l2           -     ----
      l3           7       8  }
  begin

    if rien(ss) then exit;
    fillchar(l_c,sizeof(l_c),#0);
    i:=0;niv:=1;l_element:=0;
    while i< length(ss) do
    begin
      inc(i);
      case ss[i] of
        '0'..'9','.':begin
          {on lit le montant}
          avant:=i;
          lect_nbr(ss,i);
          l_element:=i-avant+1;
          l_c[niv]:=l_c[niv]+copy(ss,avant,l_element);
          end;
        '[':begin   {nom d'un indice}
            avant:=i;
            l_c[niv]:=l_c[niv]+lect_str(ss,i,']');
            l_element:=i-avant-1;
            end;
        '@': begin   {appel à une sous chaine}
             inc(i);
             l_s:=traitement_edition(sous_chaine[trunc(lect_nbr(ss,i))]);
             l_element:=length(l_s[1])+2;
             l_c[inv(niv)]:=l_c[inv(niv)]+b(length(l_c[niv])-length(l_c[inv(niv)]))+' '+l_s[3]+' ';
             l_c[2]:=l_c[2]+b(length(l_c[niv])-length(l_c[2]))+' '+l_s[2]+' ';
             l_c[niv]:=l_c[niv]+'('+l_s[1]+')';
             end;{'@'}
         '/':begin
             l_c[2]:=cop(l_c[2],length(l_c[niv])-l_element)+car(l_element);
             l_c[inv(niv)]:=cop(l_c[inv(niv)],length(l_c[niv])-l_element);
             niv:= inv(niv);{inversion de présentation}
             end;
         '+','-':begin
             {la chaine ss n'ayant pas de parenthèse rencontrer un signe
              + ou - indique que l'on est au numérateur en effet
              dans l'exemple 1+5/3-2
              le -2 est forcément au numérateur, je considère que l'écriture
              de 1+5/-3 est non traitée
              }
             if niv=3 then niv:=1;
             l_c[1]:=l_c[1]+ss[i];
             end;
         '*':begin
             l_c[niv]:=l_c[niv]+ss[i];
             end;
      end;{case}
    end;{while i}
    result:=l_c;
  end;

  // function 6
  procedure  visualisation_expression(var pMemoEdition : TMemo);
  var l_c:schema;{les trois lignes de texte}
      i:smallint;
  begin
    for i:=max downto 1 do
      if not rien(sous_chaine[i]) then
      begin
        l_c:=traitement_edition(sous_chaine[i]);
        break;{on affiche seulement la dernière ligne du tableau}
      end;

    pMemoEdition.clear;
    for i:=1 to 3 do pMemoEdition.lines.Add(l_c[i]);
  end;

begin {TraitementFormule}

  {analyse des parenthèses et chargement des sous chaines
  correspondantes aux formules
  l'indice 1 du tableau des sous-chaines correspond à la plus
  intèrieure}
  NN:=0;
  s:=expression;
  while pos('(',s)<>0 do
  begin
    inc(NN);
    s:=suppression_parenthese_interne(s);
  end;
  inc(NN);sous_chaine[NN]:=s;

  {visualisation du chargement des sous_chaines}
  {le calcul est à faire dans l'ordre des sous_chaines}
  fillchar(val_chaine,sizeof(val_chaine),#0);
  pMemoDetail.clear;
  for i:=1 to max do if not rien(sous_chaine[i]) then
  begin
    val_chaine[i]:=calcul_chaine(sous_chaine[i]);
    pMemoDetail.lines.add(in_st(i,2)+' : '+sous_chaine[i]+' = '+re_st(val_chaine[i],15,6));
    //pMemoDetail.lines.add(re_st(val_chaine[i],15,6));
  end;
       
  {édition des sous_chaines
  le principe est de prendre la sous-chaine la moins intérieure
  on remplace tout appel à un niveau par la sous chaine correspondante
  on écrit sur 3 lignes
        les numérateurs
        trait autant de carac que présence de dénominateur
        les dénominateurs
  }
  visualisation_expression(pMemoEdition);

  pMemoEdition.Lines[0] := 'P=P°*' + pMemoEdition.Lines[0];
  pMemoEdition.Lines[1] := '     ' + pMemoEdition.Lines[1];
  pMemoEdition.Lines[2] := '     ' + pMemoEdition.Lines[2];
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. : remplace les indices de la formule par ls indices de l'affaire
                 supprime l'arrondi
Mots clefs ... :
*****************************************************************}
procedure RemplaceIndice(pStForCode, pStAffaire : String; var pStMemo : String);
Var
  St                  : String;
  Q                   : TQuery;
  TobF                : Tob;
  TobParamF           : Tob;
  i                   : Integer;
  PosVirgule          : Integer;
  PosParentheseFermee : integer ;
  StaJeter            : String;
  StIndice            : String;
  StIndiceAff         : String;

begin

  if pStAffaire <> '' then
  begin
    TobF := TOB.Create('Mes Formules',nil,-1);
    TobParamF:=TOB.Create('Mes Parametres de Formules',nil,-1);
    Try
      St:='SELECT AFE_FORCODE,AFE_FOREXPRESSION,AFE_INDCODE10, AFE_INDCODE9, AFE_INDCODE8,AFE_INDCODE7, AFE_INDCODE6,' ;
      St:=St+'  AFE_INDCODE5,AFE_INDCODE4,AFE_INDCODE3,AFE_INDCODE2,AFE_INDCODE1 ' ;
      St:=St+' FROM AFFORMULE WHERE AFE_FORCODE = "'+ pStForCode +'"';

      Q:=Nil ;
      try
        Q := OpenSQL(St, TRUE);
        TobF.LoadDetailDB('','','',Q,false);
      finally
        Ferme(Q);
      end ;

      St:='SELECT AFC_FORCODE,AFC_AFFAIRE,AFC_INDAFF10, AFC_INDAFF9, AFC_INDAFF8,AFC_INDAFF7, AFC_INDAFF6,' ;
      St:=St+'  AFC_INDAFF5,AFC_INDAFF4,AFC_INDAFF3,AFC_INDAFF2,AFC_INDAFF1 ' ;
      St:=St+' FROM AFPARAMFORMULE WHERE AFC_FORCODE = "'+ pStForCode +'" AND AFC_AFFAIRE="'+ pStAffaire +'"';
      Q:=Nil;

      try
        Q := OpenSQL(St, TRUE);
        TobParamF.LoadDetailDB('','','',Q,false) ;
      finally
        Ferme(Q) ;
      end ;

      if TobF.Detail.Count>0 then
        pStMemo:=TobF.Detail[0].getvalue('AFE_FOREXPRESSION');

      While (pos(';',pStMemo)>0) do
      begin
        PosVirgule:=pos(';',pStMemo);
        PosParentheseFermee:=pos('}',pStMemo);
        StaJeter:=copy(pStMemo,PosVirgule,PosParentheseFermee-PosVirgule+1);
        pStMemo:=Stringreplace(pStMemo,StaJeter,'', [rfReplaceAll,rfIgnoreCase]);
      end;

      pStMemo:=Stringreplace(pStMemo,'ARR{','', [rfReplaceAll,rfIgnoreCase]);

      if TobParamF.Detail.Count>0 then
        begin
        For i:=1 to 10 do
          begin
            StIndice:=trim(TobF.detail[0].getvalue('AFE_INDCODE'+inttostr(i)));
            StIndiceAff:=trim(TobParamF.detail[0].getvalue('AFC_INDAFF'+inttostr(i)));
            pStMemo:=Stringreplace(pStMemo,StIndice,StIndiceAff,[rfReplaceAll,rfIgnoreCase]);
          end;
        end;

    Finally
      TobF.free;
      TobParamF.free;
    End;
  end
  else
    pStMemo:=Stringreplace(pStMemo,'ARR{','', [rfReplaceAll,rfIgnoreCase]);

end;
                  
function FormuleEdition(pInNumLigne : Integer; pStAffaire, pStForCode : String) : String;
var
  vMemoDetail   : TMemo;
  vMemoEdition  : TMemo;
  vStFormule    : STring;
  Ecran         : TForm;

begin
                   
  Ecran := Tform.create(nil);
  vMemoDetail := TMemo.create(Ecran);
  vMemoDetail.Parent := Ecran;
  vMemoDetail.Width := 700;
  vMemoEdition := TMemo.create(Ecran);
  vMemoEdition.Parent := Ecran;
  vMemoEdition.Width := 700;
  try
    RemplaceIndice(pStForCode, pStAffaire, vStFormule);
    TraitementFormule(vStFormule, vMemoDetail, vMemoEdition);

    case pInNumLigne of
      1 : result := vMemoEdition.Lines[0];
      2 : result := vMemoEdition.Lines[1];
      3 : result := vMemoEdition.Lines[2];
    end;
  finally
    vMemoDetail.Free;
    vMemoEdition.Free;
    Ecran.free;
  end;
end;

{function FormuleEdition(pInNumLigne : Integer; pStAffaire, pStForCode : String; var pSLFormule, pSLDetail : TStringList) : Boolean;
var
  vMemoDetail   : TMemo;
  vMemoEdition  : TMemo;
  vStFormule    : STring;
  Ecran         : TForm;

begin

  Ecran := Tform.create(nil);
  vMemoDetail := TMemo.create(Ecran);
  vMemoDetail.Parent := Ecran;
  vMemoDetail.Width := 700;
  vMemoEdition := TMemo.create(Ecran);
  vMemoEdition.Parent := Ecran;
  vMemoEdition.Width := 700;
  try
    RemplaceIndice(pStForCode, pStAffaire, vStFormule);
    TraitementFormule(vStFormule, vMemoDetail, vMemoEdition);

    pSLFormule.Add(vMemoEdition.Lines[0]);
    pSLFormule.Add(vMemoEdition.Lines[1]);
    pSLFormule.Add(vMemoEdition.Lines[2]);

  finally
    vMemoDetail.Free;
    vMemoEdition.Free;
    Ecran.free;
  end;
end;
}

function FormuleDetail(pInNumLigne : Integer; pStAffaire, pStForCode : String) : String;
var
  vMemoDetail       : TMemo;
  vMemoEdition      : TMemo;
  vStFormule        : STring;
  Ecran             : TForm;
  vFormule          : TCALCULCOEF;
//  vInDernier        : Integer;
  vSlDetailEdition  : TStringList;

begin

  Ecran := Tform.create(nil);
  vMemoDetail := TMemo.create(Ecran);
  vMemoDetail.Parent := Ecran;
  vMemoDetail.Width := 700;
  vMemoEdition := TMemo.create(Ecran);
  vMemoEdition.Parent := Ecran;
  vMemoEdition.Width := 700;
  vSlDetailEdition := TStringList.create;
  vFormule := TCALCULCOEF.Create;

  try
    if (not vFormule.FormuleEditionDetail(pStAffaire, pStForCode, vStFormule)) and
      (pInNumLigne = 1) then
      result := 'La formule contient au moins un arrondi de terme, le calcul progressif n''est pas possible.'
    else
    begin
      TraitementFormule(vStFormule, vMemoDetail, vMemoEdition);
      case pInNumLigne of
        1 : result := vMemoEdition.Lines[0];
        2 : result := vMemoEdition.Lines[1];
        3 : result := vMemoEdition.Lines[2];
        4 : niveau2(vMemoEdition.Lines[0], vMemoDetail, result);
        5 : begin MemoDetailEdition(vMemoDetail, vSlDetailEdition); if vSlDetailEdition.Count > 0 then result := vSlDetailEdition.Strings[0] else result := '' end;
        6 : begin MemoDetailEdition(vMemoDetail, vSlDetailEdition); if vSlDetailEdition.Count > 1 then result := vSlDetailEdition.Strings[1] else result := ''  end;
        7 : begin MemoDetailEdition(vMemoDetail, vSlDetailEdition); if vSlDetailEdition.Count > 2 then result := vSlDetailEdition.Strings[2] else result := ''  end;
        8 : begin MemoDetailEdition(vMemoDetail, vSlDetailEdition); if vSlDetailEdition.Count > 3 then result := vSlDetailEdition.Strings[3] else result := ''  end;
      end;
    end;

  finally
    vFormule.Free;
    vMemoDetail.Free;
    vMemoEdition.Free;
    Ecran.free;
  end;
end;

// remplacement du contenu des parentheses
function niveau2(pStFormule : String; pMemo : TMemo; var pStNewFormule : String) : Integer;
var
  i       : Integer;
  d       : Integer;
  f       : Integer;
  vStDeb  : String;
  vStFin  : String;

begin
  i := 0;
  pStNewFormule := pStFormule;
  while pos('(', pStNewFormule) > 0 do
  begin
    d := pos('(',pStNewFormule);
    f := pos(')',pStNewFormule);
    vStDeb := copy(pStNewFormule, 0, d - 1);
    vStFin := copy(pStNewFormule, f + 1, length(pStNewFormule));
    pStNewFormule := vStDeb + trim(trim(copy(pMemo.lines[i], pos('=', pMemo.lines[i]) + 1, length(pMemo.lines[i])))) + vStFin;
    i := i + 1;
  end;
  result := i;
end;

procedure MemoDetailEdition(pMemoDetail : TMemo; var pSlDetailEdition : TStringList);
var
  i : Integer;
//  j : Integer;
//  k : Integer;
  s : String;
  n : Integer;

  function LaValeur(pSt : String) : String;
  var
    i : Integer;
  begin
    i := pos('=', pSt);
    result := trim(copy(pSt, i + 1, length(pSt)));
  end;

  function RemplaceValeur(pMemoDetail : TMemo; i : Integer) : String;
  var
    l : Integer;
  begin
    l := pos('=', pMemoDetail.lines[i]);
    result := copy(pMemoDetail.lines[i], 6, l-7);
    if pos('@', result) = 0 then
      result := ''
    else
      while pos('@', result) > 0 do
      begin
        l := pos('@', result);
        s := copy(result, l, 3);
        n := strToInt(copy(s, 3, 1));
        result := StringReplace(result, s, LaValeur(pMemoDetail.lines[n]), [rfReplaceAll]);
      end;
  end;

begin
//  j := 0;
  pSlDetailEdition.Add('P=P°*' + LaValeur(pMemoDetail.lines[pMemoDetail.lines.Count - 1]));
  for i := pMemoDetail.lines.Count - 1 downto 0 do
  begin
    s := RemplaceValeur(pMemoDetail, i);
    if s <> '' then
      pSlDetailEdition.insert(0, 'P=P°*' + s);
  end;
end;

{function suppression_parenthese_interne(s:string):string;
var j,dern_ouv,prem_ferm : smallint;
begin
  j:=0;dern_ouv:=0;
  while j<length(s) do
  begin
    inc(j);
    case s[j] of
      '(':begin
          dern_ouv:=j+1;
          end;
      ')':begin
          prem_ferm:=j-1;
          sous_chaine[NN]:=copy(s,dern_ouv,prem_ferm-dern_ouv+1);
          result:=cop(s,dern_ouv-2)+'@'+num_str(NN,2)+cope(s,length(s)-prem_ferm-1);
          exit;
          end;
    end;
  end;
end;
}

initialization
end.

