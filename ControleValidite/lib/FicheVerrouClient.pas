unit FicheVerrouClient;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls,DB, ADODB,AdoInt, OleDB;

type

  TEtatBase = class (TObject)
  	Nombase : string;
    IDBase : string;
    Etat : string;
  end;

  TListEtatbase = class (TList)
  	private
      function Add(AObject: TEtatBase): Integer;
      function GetItems(Indice: integer): TEtatBase;
      procedure SetItems(Indice: integer; const Value: TEtatBase);
    public
      constructor create;
    	destructor destroy; override;
      property Items [Indice : integer] : TEtatBase read GetItems write SetItems;
      procedure clear; override;
  end;

  TFverrou = class(TForm)
    LClient: TLabel;
    TCLIENT: TEdit;
    LIBCLIENT: TLabel;
    CBVERROU: TCheckBox;
    Panel1: TPanel;
    BVALIDE: TSpeedButton;
    BANNULE: TSpeedButton;
    CBDATABASE: TComboBox;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BVALIDEClick(Sender: TObject);
    procedure BANNULEClick(Sender: TObject);
  private
    { Déclarations privées }
    fClient : PChar;
    fDatabase : PChar;
    fServer : PChar;
    fListDatabases : TListEtatbase;
    procedure SetInfoTiers;
  public
    { Déclarations publiques }
    property Server : pChar read fServer write fServer;
    property client : pChar read fClient write fclient;
    property Database : PChar read fDatabase write fDatabase;
  end;


implementation
uses UAccesDatabase;

{$R *.dfm}

procedure TFverrou.FormCreate(Sender: TObject);
begin
  fListDatabases := TListEtatbase.create;
end;

procedure TFverrou.FormDestroy(Sender: TObject);
begin
  fListDatabases.free;
end;

procedure TFverrou.FormShow(Sender: TObject);
begin
	SetInfoTiers;
end;

procedure TFverrou.SetInfoTiers;
var Conn : string;
		QQ : TADOQuery;
    HH : TEtatBase;
    II : Integer;
begin
  Conn := SetConnectionString (Server,Database);
  //
  QQ := TADOQuery.create (nil);
  QQ.ConnectionString := Conn;
  QQ.SQL.Clear;
  QQ.SQL.Add('SELECT T_TIERS,T_LIBELLE FROM TIERS WHERE T_TIERS="'+fClient+'" AND T_NATUREAUXI="CLI"');
  TRY
    QQ.Open;
    TCLIENT.Text := QQ.findField('T_TIERS').AsString;
    LIBCLIENT.Caption := QQ.findField('T_LIBELLE').AsString;
  FINALLY
  	QQ.Close;
  	FreeAndNil(QQ);
	END;
  //
  QQ := TADOQuery.create (nil);
  QQ.ConnectionString := Conn;

  QQ.SQL.clear;
  QQ.SQL.Add('SELECT B01_ID,B01_LIBELLE,B01_BLOCAGE FROM BACTIVATION WHERE B01_TIERS="'+fClient+'"');
  TRY
    QQ.Open;
    while not QQ.Eof do
    begin
      HH := TEtatBase.Create;
      HH.IDBase  := QQ.Fields[0].AsString ;
      HH.Nombase := QQ.Fields[1].AsString ;
      HH.Etat    := QQ.Fields[2].AsString ;
      fListDatabases.Add(HH);
    end;
  finally
    QQ.Close;
    QQ.Free;
  end;
  //
  CBDATABASE.Clear;
  For II := 0 to fListDatabases.Count -1 do
  begin
  	CBDATABASE.AddItem(fListDatabases.Items [II].Nombase,nil);
  end;
  //
end;

{ TListEtatbase }

function TListEtatbase.Add(AObject: TEtatBase): Integer;
begin
	Result := inherited Add(Aobject);
end;

procedure TListEtatbase.clear;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := count -1 downto 0 do
    begin
      if TEtatBase(Items [Indice])<> nil then
      begin
         TEtatBase (Items [Indice]).free;
         Items[Indice] := nil;
      end;
    end;
  end;
  inherited;
end;

constructor TListEtatbase.create;
begin

end;

destructor TListEtatbase.destroy;
begin
  clear;
  inherited;
end;

function TListEtatbase.GetItems(Indice: integer): TEtatBase;
begin
  result := TEtatBase (Inherited Items[Indice]);
end;

procedure TListEtatbase.SetItems(Indice: integer; const Value: TEtatBase);
begin
  Inherited Items[Indice]:= Value;
end;

procedure TFverrou.BVALIDEClick(Sender: TObject);
begin
//	close;
end;

procedure TFverrou.BANNULEClick(Sender: TObject);
begin
	close;
end;

end.
