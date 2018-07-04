unit ySouche_abs;

// Déclaration de l'interface des souches

interface

type
  IySouche = interface
  ['{AAEAFF33-372E-4C94-889A-D6A77523CB35}']
    procedure CreateFirst(const Prefixe: String; const Contexte: String; iNumero : integer);
    function ReadCurrent(const Prefixe: String; const Contexte: String) : integer;
    function GetNext(const Prefixe: String; const Contexte: String; Combien : integer = 1): Integer;
    function Exists(const Prefixe, Contexte: String): Boolean;
    procedure SetValue (const Prefixe, Contexte : string; Compteur : integer);
    procedure Remove (const Prefixe, Contexte : string);
  end;

implementation

end.
