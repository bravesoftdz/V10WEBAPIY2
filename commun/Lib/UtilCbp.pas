{***********UNITE*************************************************
Auteur  ...... : MD3
Créé le ...... : 19/05/2009
Modifié le ... : 19/05/2009
Description .. : Unité d'encapsulation des fonctions deprecated CBP 7.4
Suite ........ : A utiliser au lieu de UtilPGI dans lequel ces fonctions seront
Suite ........ : supprimées.
Suite ........ : M.Desgoutte/D.Brosset
Mots clefs ... :
*****************************************************************}
unit UtilCbp;


interface

{$IFNDEF EAGLSERVER}
uses
   Hdb
  ,Controls
  ,HTB97
  ,forms
  ,HCtrls ;
{$ENDIF !EAGLSERVER}

function ChampToLibelle_ (Const NomChamp : string): string;
function ChampToType_ (Const NomChamp : string) : string;
function ChampToNum_ (Const NomChamp : string) : integer;
function PrefixeToNum_ (Const Prefixe : string) : integer;
function PrefixeToTable_ (Const Prefixe : string) : string;


function TableToNum_ (Const NomTable : string) : integer;
function TableToCle1_ (Const NomTable : string): string;
function TableToLibelle_ (Const NomTable : string): string;

function usDateTime_ (Const laDate : TDateTime) : string;
function UsDateTimeToDateTime_(Const laDate : String) : tDateTime;
function usTime_ (Const laDate : TDateTime) : string;
function GetRTFStringText (lestring : HString) : Widestring;


implementation

uses
  HRichOLE // pour ChampToTT (manquante dans CbpMCD)
  ;


function GetRTFStringText (LeString : HString) : WideString;
var fTexte : ThRichEditOle;
		fInternalWindow :TToolWindow97;
begin
  fInternalWindow := TToolWindow97.create(Application.MainForm);
  fInternalWindow.Parent := Application.MainForm;
  fInternalWindow.Visible := false;
  fInternalWindow.Width := 600;
  //
  fTexte := THRichEditOLE.Create (fInternalWindow);
  ftexte.Parent := fInternalWindow;
  ftexte.text := '';
  ftexte.Align := alClient;
  TRY
 		StringToRich(ftexte,lestring);
  	result := fTexte.lines.Text;
  finally
  	fTexte.Free;
    fInternalWindow.Free;
  end;
end;

function TableToCle1_ (Const NomTable : string): string;
begin
	Result := TableToCle1(NomTable);
end;

function TableToLibelle_ (Const NomTable : string): string;
begin
	Result := TableToLibelle(NomTable);
end;


function ChampToLibelle_ (Const NomChamp : string): string;
begin
  result := ChampToLibelle(NomChamp);
end;

function PrefixeToNum_ (Const Prefixe : string) : integer;
begin
	Result := PrefixeToNum(prefixe);
end;

function PrefixeToTable_ (Const Prefixe : string) : string;
begin
	Result := PrefixeToTable(Prefixe);
end;

function ChampToType_ (Const NomChamp : string) : string;
begin
	Result := ChampToType(NomChamp) ;
end;


function ChampToNum_ (Const NomChamp : string) : integer;
begin
	Result := ChampToNum(NomChamp) ;
end;

function TableToNum_ (const NomTable : string) : integer;
begin
  Result := TableToNum(NomTable);
end;

function usDateTime_ (Const laDate : TDateTime) : string;
begin
  Result := USDATETIME(laDate);
end;

function UsDateTimeToDateTime_(Const laDate : String) : tDateTime;
begin
  Result := UsDateTimeToDateTime (laDate);
end;

function usTime_ (Const laDate : TDateTime) : string;
begin
  Result := USTIME(LaDate);
end;

end.




