unit YSouche;

// implementation de l'interface des souches

interface

uses
  ySouche_abs;

function yISouche : IySouche;

implementation

uses
  uTob,
  {$IFNDEF EAGLCLIENT}
    uDbxDataSet,
  {$ENDIF}
  HCtrls,
  cbpDBSequences,
  gSysUtils, //wCommuns,
  Hent1,
  SysUtils, DB;

const
  ySoucheTableName = 'SOUCHE';
  YSouche_SequenceCode_Separator = '~';
  ySoucheCtx = 'SH' + ySouche_SequenceCode_Separator;

type
  { gestion des jetons façon "séquences CBP" ----------------------------------------------------- }
  TySouche = class(TInterfacedObject, IySouche)
  private
    function GetCode (const Prefixe, Contexte: String) : String;
  public
    procedure CreateFirst (const Prefixe: String; const Contexte: String; iNumero : integer);
    function ReadCurrent(const Prefixe: String; const Contexte: String) : integer;
    function GetNext(const Prefixe: String; const Contexte: String; Combien : integer = 1): Integer;
    function Exists(const Prefixe, Contexte: String): Boolean;
    procedure SetValue (const Prefixe, Contexte : string; Compteur : integer);
    procedure Remove (const Prefixe, Contexte : string);
  private
    function CreateIfNotExist (const Prefixe, Contexte: String; var Numero : integer; Combien : integer = 0) : boolean;
  end;

function yISouche : IySouche;
begin
  Result := TySouche.Create() as IySouche
end;

function TySouche.GetCode(const prefixe, contexte : String): String;
begin
  Result := ySoucheCtx
          + Trim(TGStrUtils.PadRight(prefixe, 3, ' ')
          + Trim(iif(Contexte <> '', ySouche_SequenceCode_Separator  + contexte, '')))
end;

procedure TySouche.CreateFirst(const Prefixe: String; const Contexte: String; iNumero : integer);
begin
  DBSequence.CreateSequence(GetCode(Prefixe, Contexte), iNumero)
end;

function TySouche.Exists(const Prefixe, Contexte: String): Boolean;
begin
  Result := DBSequence.ExistSequence(GetCode(Prefixe, Contexte))
end;

{ renvoie le dernier n° de séquence pris }
function TySouche.ReadCurrent(const Prefixe, Contexte: String) : integer;
begin
  Result := 0;
  try
    Result := DBSequence.ReadCurrent(GetCode(Prefixe, Contexte))
  except
    { séquence inexistante : on la crée à 0 en tant qque n° en cours }
    if not CreateIfNotExist (Prefixe, Contexte, Result) then raise;
  end;
end;

{ renvoie le prochain n° de séquence théorique }
function TySouche.GetNext(const Prefixe, Contexte: String; Combien : integer = 1): Integer;
begin
  Result := 0;
  try
    Result := DBSequence.GetNext (GetCode(Prefixe, Contexte), Combien)
  except
    { séquence inexistante : on la crée au numéro de la souche tant que n° en cours }
    if not CreateIfNotExist (Prefixe, Contexte, Result, Combien) then
    begin
      V_PGI.IoError := oeUnknown;
      Raise;
    end;
  end
end;

function TySouche.CreateIfNotExist (const Prefixe, Contexte: String; var Numero : integer; Combien : integer = 0): boolean;
var
  QNum : TQuery;
begin
  Result := false;
  if not Exists(Prefixe, Contexte) then
  begin
    QNum := OpenSql ('SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="' + Prefixe + '" AND SH_SOUCHE="' + Contexte + '"', true);
    try
      Numero := QNum.Fields [0].AsInteger + Combien;
      CreateFirst(Prefixe, Contexte, Numero);
      Result := true;
    finally
      Ferme (QNum);
    end;
  end;
end;

procedure TySouche.SetValue (const Prefixe, Contexte : string; Compteur : integer);
begin
  try
    if (Compteur = 0) or not Exists (Prefixe, Contexte) or (Compteur > ReadCurrent (Prefixe, Contexte)) then
      CreateFirst (Prefixe, Contexte, Compteur);
  except
    on e : Exception do
    begin
      raise(exception.create(e.Message+ ' - Souche : ' + Contexte + ' - Compteur : ' + IntToStr (compteur)));
    end;
  end;
end;

procedure TySouche.Remove (const Prefixe, Contexte : string);
begin
  try
    DBSequence.RemoveSequence (GetCode(Prefixe, Contexte));
  except
    on e : Exception do
    begin
      raise(exception.create(e.Message+ ' - Souche : ' + Contexte));
    end;
  end;
end;

end.
