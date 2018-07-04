unit uEnreg;

interface

uses classes;

type

  noChamp = integer;

  Tenreg = class(Tobject)
    // private
    //        list: TstringList;
  private
  public
    list: TstringList;
    constructor create();
    destructor destroy(); override;
    function ch(N: noChamp): string;
    procedure assign(Source: Tenreg);
    function nbChamp(): integer;
  end;

implementation

constructor Tenreg.create();
begin
  inherited create();
  list := TstringList.create;
end;

destructor Tenreg.destroy();
begin
  list.free;
  inherited destroy;
end;

procedure Tenreg.assign(Source: Tenreg);
begin
  list.assign(Source.list);
end;

function Tenreg.ch(N: noChamp): string;
begin
  Result := list[N - 1]; // list commence à zéro mais les champ commence à 1
end;

function Tenreg.nbChamp(): integer;
begin
  Result := list.count;
end;

end.
