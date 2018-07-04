unit UEncryptage;


interface

uses classes,SysUtils,IdHashMessageDigest, idHash;

const
	decalage = 70;

function crypte(text:string):string;          //Fonction pour crypter la chaine
function decrypte(text:string):string;       //Fonction pour décrypter la chaine
function MD5(const fileName : string) : string;
function readtokenST (var TheChaine : string; Separator : string) : string;

implementation

function readtokenST (var TheChaine : string; Separator : string) : string;
var II : integer;
begin
  II:= Pos(Separator,TheChaine);
  if II > 0 then
  begin
    Result := Copy(TheChaine,1,II-1);
    TheChaine := Copy(TheChaine,II+1,Length(TheChaine)-II);
  end;
end;

function crypte(text:string):string;          //Fonction pour crypter la chaine
var pos:integer;
		text1:string;
begin
	result := '';
	if text = '' then exit;
  text1 := text;
  for pos := 1 to length(text1) do
  	text1[pos] := chr(ord(text1[pos]) + decalage);        //crypte la chaine
  crypte := text1;
end;

function decrypte(text:string):string;       //Fonction pour décrypter la chaine
var pos:integer;
  	text1:string;
begin
	result := '';
	if text = '' then exit;
  text1 := text;
  for pos := 1 to length(text1) do
  	text1[pos] := chr(ord(text1[pos]) - decalage);             //decrypte la chaine
  result := text1;
end;

function MD5(const fileName : string) : string;
 var
   idmd5 : TIdHashMessageDigest5;
   fs : TFileStream;
//   hash : T4x4LongWordRecord;
 begin
   idmd5 := TIdHashMessageDigest5.Create;
   fs := TFileStream.Create(fileName, fmOpenRead OR fmShareDenyWrite) ;
   try
     result := idmd5.AsHex(idmd5.HashValue(fs)) ;
   finally
     fs.Free;
     idmd5.Free;
   end;
 end;
end.




