unit ULList;
// LITE Tlist : c'est une Tlist avec taille pré-initialisé
//             mais qui fait 8 octet de moins qu'un Tlist

{$R-}

interface

uses SysUtils, Windows;

const

  { Maximum TList size }

  MaxListSize = Maxint div 16;

type
  { TLList class }
  {  ATTENTION pour détruire cette liste il faut appeller
        L.clear(Size);
        L.free;
  }
  PPointerList = ^TPointerList;
  TPointerList = array[0..MaxListSize - 1] of Pointer;

  TLList = class(TObject)
  private
    FList: PPointerList;

  protected
    procedure SetCapacity(OldCap, NewCapacity: Integer);
    procedure Put(Index: Integer; Item: Pointer);
    function Get(Index: Integer): Pointer;
  public
    constructor create(SizeList: integer);
    destructor Destroy; override;
    procedure Clear();
    property Items[Index: Integer]: Pointer read Get write Put; default;
    property List: PPointerList read FList;
  end;

implementation

constructor TLList.create(SizeList: integer);
begin
  SetCapacity(0, SizeList);
end;

destructor TLList.destroy();
begin
  clear();
end;

procedure TLList.Clear();
begin
  ReallocMem(FList, 0);
end;

function TLList.Get(Index: Integer): Pointer;
begin
  Result := FList^[Index];
end;

procedure TLList.Put(Index: Integer; Item: Pointer);
begin
  FList^[Index] := Item;
end;

procedure TLList.SetCapacity(OldCap, NewCapacity: Integer);
begin
  if NewCapacity <> OldCap then
  begin
    ReallocMem(FList, NewCapacity * SizeOf(Pointer));
  end;
end;

end.
