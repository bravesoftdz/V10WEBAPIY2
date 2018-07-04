unit UAuditPerf;

interface
Uses Classes,
		 Forms,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$else}
{$ENDIF}
     uTob,
     Hctrls,
     sysutils,
     DateUtils,
     UtilsRapport,
     HEnt1;
type
  TAuditItem = class;

  TlistAuditItem = class (TList)
    private
      fLevel : Integer;
      fCurrentItem : TAuditItem;
      procedure SetItems(Indice: integer; const Value: TAuditItem);
    	function GetItems(Indice: integer): TAuditItem;
    public
    	destructor destroy; override;
    	property Items [Indice : integer] : TAuditItem read GetItems write SetItems;
      property CurrentItem : TAuditItem read fCurrentItem write fCurrentItem;
      property level : Integer read fLevel write fLevel;
    	function SetDebut(Nom: string; DDebut: TdateTime): Integer;
      function Add(Item: TAuditItem): Integer;
      function SetFin (Nom : string; DFin : TDateTime) : boolean;

      procedure init;
  end;

  TAuditItem = Class (TObject)
  	private
      fDateDebut : TDateTime;
      fDateFin   : TDateTime;
      fname : string;
      ffille : TlistAuditItem;
      flevel : Integer;
    public
      constructor create;
      destructor destroy; override;
      property Nom : string read fname write fName;
      property Debut : TDateTime read fDateDebut write fDateDebut;
      property Fin : TDateTime read fDateFin write fDateFin;
      property fille : TlistAuditItem read ffille write ffille;
      property niveau : Integer read flevel write flevel;
  end;

  TAuditClass = class (TObject)
  private
    fListItems : TlistAuditItem;
    Rapport : TGestionRapport;
  public
		constructor create (FF : TForm);
    destructor destroy; override;
    procedure Debut (Nom : string);
    procedure Fin (Nom : string);
    procedure AfficheAudit;
    procedure ReInit;
  end;

implementation

{ TAuditClass }

procedure AddLigneRapport (Rapport : TGestionRapport;TheItem: TAuditItem; TextePrec : string);
var TheText,TextSuiv : string;
		indice,hh,mm,ss : Integer;
    nextItem : TAuditItem;
    Fill : string;
    PreString : string;
    D2 : TDateTime;
begin
  fill := StringOfChar(chr(9),30);
  D2 := TheItem.Fin;
  hh := DaysBetween(TheItem.Debut,D2);
  D2 := IncDay(D2,(hh*-1));
  mm := MinutesBetween(TheItem.Debut,D2);
  D2 := IncMinute(D2,(mm*-1));
  Ss := SecondsBetween(TheItem.Debut,D2);

	if TheItem.fille.Count = 0 then
  begin
    Prestring := Copy(Fill,1,TheItem.flevel);
    TheText := prestring+TextePrec + ' ' + TheItem.nom + '      DEBUT : '+TimeToStr(TheItem.Debut)+
    					' FIN : '+TimeToStr(TheItem.Fin)+ ' --> DELAI : '+inttostr(hh)+'h'+IntToStr(mm)+'mn'+IntToStr(ss)+'sec';
    Rapport.SauveLigMemo (TheText);
  end else
  begin
    Prestring := Copy(Fill,1,TheItem.flevel);
    TheText := PreString+TextePrec + ' ' + TheItem.nom + ' DEBUT : '+TimeToStr(TheItem.Debut);
    Rapport.SauveLigMemo (TheText);
    for Indice := 0 to TheItem.Fille.Count -1 do
    begin
      nextItem := TheItem.Fille.items[Indice];
			TextSuiv := texteprec+'.'+InttoStr(Indice+1);
      AddLigneRapport (Rapport,nextItem,TextSuiv);
    end;
    TheText := prestring+TextePrec + ' ' + TheItem.nom + ' FIN : '+TimeToStr(TheItem.Fin)+
    					' --> DELAI : '+inttostr(hh)+'h'+IntToStr(mm)+'mn'+IntToStr(ss)+'sec';
    Rapport.SauveLigMemo (TheText);
  end;
end;

procedure TAuditClass.AfficheAudit;
var Indice : Integer;
    TheItem : TAuditItem;
begin
  for Indice := 0 to fListItems.Count -1 do
  begin
    TheItem := fListItems.Items [Indice];
    AddLigneRapport (Rapport,TheItem, IntToStr(Indice+1));
  end;
  Rapport.AfficheRapport;
  Reinit;
end;

constructor TAuditClass.create (FF : TForm);
begin
  fListItems := TlistAuditItem.Create;
  fListItems.level := 0;
  //
  Rapport := TGestionRapport.Create (FF);
  Rapport.Affiche := True;
  Rapport.titre := 'Audit des systèmes';
  Rapport.Print := false;
  Rapport.Close := True;
  Rapport.Sauve := false;
end;

procedure TAuditClass.Debut(Nom: string);
begin
	fListItems.SetDebut(Nom,Now);
end;

destructor TAuditClass.destroy;
begin
  Rapport.Free;
  fListItems.free;
  inherited;
end;


procedure TAuditClass.Fin(Nom: string);
var TheItem : TAuditItem;
begin
  fListItems.SetFin(Nom,Now);
end;

procedure TAuditClass.ReInit;
begin
	fListItems.init;
end;

{ TAuditItem }

constructor TAuditItem.create;
begin
  fDateDebut := now;
  fDateFin   := now;
  ffille := TlistAuditItem.Create;
  fname := '';
end;

destructor TAuditItem.destroy;
begin
	ffille.Free;
  inherited;
end;



{ TlistAuditItem }

function TlistAuditItem.SetDebut(Nom : string; DDebut : TdateTime): Integer;
var Item : TAuditItem;
begin
  if fCurrentItem = nil then
  begin
    Item := TAuditItem.create;
    Item.flevel := fLevel;
    Item.fille.level := fLevel + 1;
    Item.Nom := Nom;
    Item.Debut := now;
  	result := Add(Item);
    fCurrentItem := Item;
  end else
  begin
  	result := fCurrentItem.fille.SetDebut(Nom,DDebut);
  end;
end;


destructor TlistAuditItem.destroy;
var indice : Integer;
begin
  init;
  inherited;
end;

function TlistAuditItem.GetItems(Indice: integer): TAuditItem;
begin
  result := TAuditItem (Inherited Items[Indice]);
end;

procedure TlistAuditItem.init;
var indice : Integer;
begin
  if count > 0 then
  begin
    for Indice := count -1 downto 0 do
    begin
      if TAuditItem(Items [Indice])<> nil then TAuditItem(Items [Indice]).free;
    end;
  end;
  clear;
end;

procedure TlistAuditItem.SetItems(Indice: integer;const Value: TAuditItem);
begin
  Inherited Items[Indice]:= Value;
end;

function TlistAuditItem.Add(Item: TAuditItem): Integer;
begin
	Result := inherited Add(Item);
end;

function TlistAuditItem.SetFin(Nom: string; DFin: TDateTime) : boolean;
var Indice : Integer;
begin
	Result := false;
	if fCurrentItem.Nom = Nom then
  begin
		fCurrentItem.Fin := DFin;
    fCurrentItem := nil;
    Result := True;
  end else
  begin
    for Indice := 0 to fCurrentItem.fille.Count -1 do
    begin
    	Result :=  fCurrentItem.fille.SetFin(Nom,DFin);
      if Result then break;
    end;
  end;
end;

end.
