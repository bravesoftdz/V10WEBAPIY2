unit UScriptDelim;

interface

uses Classes, DBTables, Dialogs, DB, Forms,HCtrls,Sysutils;

type
        TChampListDelim = class;
        TChampClass = class of TChampParam;

        TChampParam = class(TCollectionItem)
        protected
          FName       : string[64];
          FLongueur   : Smallint;
          FVisible    : Boolean;
          FSeparateur : Boolean;
          FAligne     : Boolean;
          procedure SetName(AValue: string);
          function GetName : string;

        public
          constructor Create(ACollection: TCollection); override;
          destructor Destroy; override;

          procedure Assign(AChamp: TChampParam);
       published (* csLoading *)
          property Name: string read GetName write SetName;
          property Longueur: Smallint read FLongueur write FLongueur;
          property Visible: Boolean read FVisible write FVisible;
          property Separateur: Boolean read FSeparateur write FSeparateur;
          property Aligne: Boolean read FAligne write FAligne;

        end;

        (* ------------------------------------------------------------------ *)
        (*                       T C H A M P L I S T                          *)
        (* ------------------------------------------------------------------ *)

        TChampListDelim = class(TCollection)
        private
          function Get(itemIndex: Integer): TChampParam;
          procedure Put(index: integer; AValue: TChampParam);
        protected

        public
          destructor Destroy; override;

          procedure Assign(AList: TChampListDelim);

          procedure AddObject(AName: string; AChamp: TChampParam);
          procedure Delete(Index: Integer);
          function IndexOf(AValue: string; deb: integer=0): integer;
          procedure Move(curIndex, NewIndex: Integer);

          property Items[Index: Integer]: TChampParam read Get write Put; default;
        end;


	TScriptDelimite = class(TComponent)
	private
                FName        : string[35];
		FNomFichier  : String;
		FNomSortie   : String;
                Fdomaine     : string;
                Flibelle     : string;
                FShellexec   : string;
                FChainref     : string;
                FLongref     : integer;
                FLondeb      : integer;
                FLonfin      : integer;
                FChainelim   : string;
                FCbdelim     : string;
                DelimLibelle : string;

                FChamps      : TChampListDelim;
                procedure SetName(AValue: string);
                function GetName : string;
	public
		constructor Create(AOwner : TComponent); override;
		destructor Destroy; override;
                procedure Assign(AScript: TScriptDelimite);
                procedure SaveDelimTo(AStream: TStream);

	published
                property Name       : string read GetName write SetName;
		property NomFichier : String read FNomFichier write FNomFichier;
		property NomSortie  : String read FNomSortie write FNomSortie;
		property Domaine    : String read FDomaine write FDomaine;
		property Shellexec  : String read FShellexec write FShellexec;
		property Libelle    : String read FLibelle write FLibelle;
		property Chainref   : String read FChainref write FChainref;
		property Chainelim  : String read FChainelim write FChainelim;
                property Longref    : integer read FLongref write FLongref default 0;
                property Londeb     : integer read FLondeb write FLondeb default 0;
                property Lonfin     : integer read FLonfin write FLonfin default 0;
                property Champ      : TChampListDelim read FChamps write FChamps;
		property Cbdelim    : String read FCbdelim write FCbdelim;
                property Delimlib   : String read DelimLibelle write DelimLibelle;
	end;

function LoadScriptFromStreamDelim(AStream: TStream): TScriptDelimite;

implementation


type
  TMonReader = class(TReader)
  protected
    function Error(const Message: string): Boolean; override;
  end;

var
  ErrorReader: TStringList; // repertorie toutes les erreurs du TMonReader

function TMonReader.Error(const Message: string): Boolean;
begin
  if not Assigned(ErrorReader) then
    ErrorReader := TStringlist.Create;
  ErrorReader.Add(Message);
  Result := True;
end;

(* ------------------------------------------------------------------ *)
(*                      T S C R I P TDelimite                                 *)
(* ------------------------------------------------------------------ *)

constructor TScriptDelimite.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FChamps := TChampListDelim.Create(TChampParam);
end;


destructor TScriptDelimite.Destroy;
begin
  FChamps.Free;
end;


procedure TScriptDelimite.Assign(AScript: TScriptDelimite);
var
  N       : Integer;
  AChamp  : TChampParam;
begin

  FName       := AScript.FName;
  FNomFichier := AScript.FNomFichier;
  Fdomaine    := AScript.Fdomaine;
  Flibelle    := AScript.Flibelle;


  for N := 0 to AScript.FChamps.Count - 1 do
  begin
            AChamp := TChampParam.Create(FChamps);
            AChamp.Assign(AScript.FChamps[N]);
  end;
end;


function TScriptDelimite.GetName: string;
begin
  Result := FName;
end;

procedure TScriptDelimite.SetName(AValue: string);
begin
  FName := AValue;
end;


procedure TScriptDelimite.SaveDelimTo(AStream: TStream);
var
  ST1, ST2: TStream;
  WR: TWriter;
begin (* SaveTo *)
  ST1 := TMemoryStream.Create;
  WR := TWriter.Create(ST1, 4096);
  try
    WR.WriteRootComponent(Self);
  except
    WR.Free;
    ST1.Free;
    ShowMessage(Exception(ExceptObject).message);
    Exit;
  end;
  WR.Free;

  ST1.Seek(0, 0);
  ST2 := AStream;
  try
    ObjectBinaryToText(ST1, ST2);
  finally
    ST1.Free;
  end;
end;

function LoadScriptFromStreamDelim(AStream: TStream): TScriptDelimite;
var
  AScript: TScriptDelimite;
  RD: TMonReader;
  ST1, ST2: TStream;
begin
  ST1 := AStream;
  ST2 := TMemoryStream.Create;
  try
    ObjectTextToBinary(ST1, ST2);
  except
    ST2.Free;
    raise;
  end;
  ST2.Seek(0, soFromBeginning);
  RD := TMonReader.Create(ST2, 4096);
  try
    AScript := RD.ReadRootComponent(nil) as TScriptDelimite;
  except
    RD.Free;
    ST2.Free;
    raise;
  end;
  RD.Free;
  ST2.Free;

  Result := AScript;
end;
(* ------------------------------------------------------------------ *)
(* TChampListDelim                                                         *)
(* ------------------------------------------------------------------ *)

destructor TChampListDelim.Destroy;
begin
  inherited Destroy;
end;

procedure TChampListDelim.Assign(AList: TChampListDelim);
var
  N     : Integer;
  AChamp: TChampParam;
begin
  for N := 0 to AList.Count - 1 do
  begin
    AChamp := TChampParam.Create(self);
    AChamp.Assign(AList[N]);
  end;
end;

function TChampListDelim.Get(itemIndex: integer): TChampParam;
begin
  Result := TChampParam(inherited Items[itemIndex]);
end;

procedure TChampListDelim.Put(index: integer; AValue: TChampParam);
begin
  Items[index].Assign(AValue);
end;

function TChampListDelim.IndexOf(AValue: string; deb: integer=0): integer;
var
  N, i : integer;
  NN   : string;
begin
  Result := -1;
  for N := 0 to Count - 1 do
    if AnsiCompareText(TChampParam(Items[N]).FName, AValue) = 0 then
    begin
         Result := N;
         exit;
    end;
  if Result = -1 then
  begin
  for N := deb to Count - 1 do
  begin
    i := pos('_', TChampParam(Items[N]).FName);
    if i <> 0 then NN := copy(TChampParam(Items[N]).FName, 1, i-1);
    if AnsiCompareText(NN, AValue) = 0  then
    begin
         result := N;
         exit;
    end;
  end;

  end;
end;

procedure TChampListDelim.Move(curIndex, NewIndex: Integer);
begin
  Items[curIndex].Index := NewIndex;
end;

procedure TChampListDelim.Delete(Index: Integer);
begin { New Dispose}
  TChampParam(Items[Index]).Free;
end;

procedure TChampListDelim.AddObject(AName: string; AChamp: TChampParam);
begin
  Add.Assign(AChamp);
end;


constructor TChampParam.Create(ACollection: TCollection);
begin
  inherited Create(ACollection);
end;

destructor TChampParam.Destroy;
begin
  inherited Destroy;
end;

procedure TChampParam.Assign(AChamp: TChampParam);
begin
  FName := AChamp.FName;
  FLongueur := AChamp.FLongueur;
  FVisible := AChamp.FVisible;
  FSeparateur := AChamp.FSeparateur;
  FAligne := AChamp.FAligne;
end;

procedure TChampParam.SetName(AValue: string);
begin
  FName := AValue;
end;

function TChampParam.GetName: string;
begin
  Result := FName;
end;


initialization
  RegisterClasses([TScriptDelimite, TChampParam, TChampListDelim]);

end.
