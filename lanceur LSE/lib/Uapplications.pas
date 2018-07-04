unit Uapplications;

interface
uses ShellApi, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
		 Dialogs, StdCtrls, ExtCtrls, Registry,CBPPath,hent1;

type

	TConnection = class (TObject)
  	public
      LastSoc : Integer;
      LastUser : string;
  end;

  TappliLSE = class (TObject)
  	private
      fexe : string;
      fTitre : string;
      fBitmap : TBitmap;
      fAccepteparam : Boolean;
      //
    	procedure SetExe(const Value: string);
    	procedure TIconToTBitmap(FIcon: TIcon; var FBitmap: TBitmap);
    	procedure GetAssociatedBmp(const Ext: PChar ; var BM : Tbitmap);
    public
      constructor create;
      destructor destroy; override;
      property AppliLSE : string read fexe write SetExe;
      property BITMap : TBitmap read fBitmap;
      property Titre : string read fTitre;
      property AccepteParam : Boolean read fAccepteparam write fAccepteparam;
  end;

  TlistAppliLse = class (TList)
  	private
      function Add(AObject: TappliLSE): Integer;
      function GetItems(Indice: integer): TappliLSE;
      procedure SetItems(Indice: integer; const Value: TappliLSE);
    public
    	destructor destroy; override;
      property Items [Indice : integer] : TappliLSE read GetItems write SetItems;
    	procedure LanceAppli(Base, user, password: string; Item: Integer);
  end;

procedure ConstitueListeAppli( ListeAppli : TlistApplilse);
procedure GetLastSocSelected ( var RConnexion : TConnection);
procedure SetLastSocSelected ( RConnexion : TConnection);

implementation

procedure SetLastSocSelected ( RConnexion : TConnection);
var iReg : TregInifile;
begin
  Ireg := TRegIniFile.Create;
  iReg.WriteInteger('SOFTWARE\LSE\Lanceur','LastSoc',RConnexion.LastSoc);
  iReg.WriteString('SOFTWARE\LSE\Lanceur','LastUser',RConnexion.LastUser);
  FreeAndNil(iReg);
end;

procedure GetLastSocSelected ( var RConnexion : TConnection);
var iReg : TregInifile;
begin
  RConnexion.LastSoc := 0;
  RConnexion.LastUser := '';
  Ireg := TRegIniFile.Create;
  IReg.RootKey := HKEY_CURRENT_USER;
  RConnexion.LastSoc := IReg.ReadInteger('SOFTWARE\LSE\Lanceur','LastSoc',0);
  iReg.WriteInteger('SOFTWARE\LSE\Lanceur','LastSoc',RConnexion.LastSoc);
  RConnexion.LastUser := IReg.ReadString('SOFTWARE\LSE\Lanceur','LastUser','');
  iReg.WriteString('SOFTWARE\LSE\Lanceur','LastUser',RConnexion.LastUser);
  FreeAndNil(iReg);
end;

procedure GetInfosFromMenu(TheRepert,Menu : string; var Titre : string; var AccepteParam: boolean);
var F: TextFile;
		S: string;
    Fichier : string;
begin
  AccepteParam := true;
  Fichier := IncludeTrailingPathDelimiter(TheRepert)+Menu;
	AssignFile(F, Fichier);
  Reset(F);
  While Not Eof(F) do
  begin
    Readln(F,S);
    if Copy(S,1,6)='TITRE=' then
    begin
			Titre := Copy(S,7,255);
    end else if Copy(S,1,10)='NOPGIEXEC=' then
    begin
			AccepteParam := (Copy(S,11,5)='FALSE');
    end;
  end;
  CloseFile(F);
end;


function AjouteUneApplication ( TheList : TlistAppliLse; TheRepert,menu : string; Var MaxTaillelib : integer ) : Integer;
var OneAppli : TappliLSE;
		Applilse : string;
    Titre : string;
    Accepteparam : Boolean;
begin
  Result := -1;
  if TheList = nil then Exit;
  Applilse := IncludeTrailingPathDelimiter(TheRepert)+ ChangeFileExt(Menu,'.exe');
  if Not FileExists(AppliLSe) then exit;
  OneAppli := TappliLSE.create;
  OneAppli.AppliLSE  := AppliLse;
  getInfosFromMenu(TheRepert,Menu,Titre,AccepteParam);
  OneAppli.fTitre := Titre;
  if Length(Titre)> MaxTaillelib then MaxTaillelib := Length(Titre);
  OneAppli.AccepteParam := Accepteparam;
//  if Menu <> 'SQLADM.mnu' then OneAppli.AccepteParam := Accepteparam
//  												else OneAppli.AccepteParam := false;
  Result := TheList.add(OneAppli);
end;

procedure AjouteApplications (ListeAppli : TlistApplilse; Repertoire : string; Var MaxTailleLib : integer) ;
var TheRepert,NomFic : string;
		Rec : TSearchRec;
begin
	TheRepert := IncludeTrailingPathDelimiter(Repertoire)+'APP';
  if not DirectoryExists(TheRepert) then Exit;
  NomFic := IncludeTrailingPathDelimiter(TheRepert)+'*.MNU';
  if FindFirst (Nomfic,faAnyFile,Rec) = 0 then
  begin
    AjouteUneApplication (ListeAppli,TheRepert,rec.name, MaxTailleLib);
    while FindNext (Rec) = 0 do
    begin
    	AjouteUneApplication (ListeAppli,TheRepert,rec.name,MaxTailleLib);
    end;
  end;
  FindClose (Rec);
end;

procedure AjusteLibelles(ListeAppli : TlistApplilse;  MaxTailleLib : integer);
var i : Integer;
    APPLSE : TappliLSE;
begin
	For i := 0 to ListeAppli.Count -1 do
  begin
		APPLSE := ListeAppli.GetItems(I);
    if Length(APPLSE.fTitre) < MaxTailleLib then
    begin
      APPLSE.fTitre := StringOfChar(' ',MaxTailleLib-length(APPLSE.fTitre)) + APPLSE.fTitre ;
    end;
  end;
end;

procedure ConstitueListeAppli( ListeAppli : TlistApplilse);
var Repert : string;
	 iReg : TregInifile;
   iPos : Integer;
   MaxTailleLib : Integer;
begin
  MaxTailleLib := 0;
  Ireg := TRegIniFile.Create;
  IReg.RootKey := HKEY_LOCAL_MACHINE;
  repert := IncludeTrailingPathDelimiter(IReg.ReadString('SOFTWARE\CEGID\Cegid Business','INSTALLDIR',IncludeTrailingPathDelimiter(UpperCase(TCbpPath.GetCegid))+'Cegid Business'));
  FreeAndNil(iReg);
  iPos := Pos('\CEGID\',Repert);
  Repert := Copy(Repert,1,iPos+6);
  AjouteApplications(ListeAppli,repert+'Cegid Business',MaxTailleLib);
  AjouteApplications(ListeAppli,repert+'LSE Business',MaxTailleLib);
  AjouteApplications(ListeAppli,repert+'Cegid Common',MaxTailleLib);
  // AjusteLibelles(ListeAppli,MaxTailleLib);
end;


function ISAccepteParam (NomMenu : string) : Boolean;
begin
  Result := false; 
end;

{ TappliLSE }

constructor TappliLSE.create;
begin
  fexe  := '';
  fTitre := '';
  fBitmap := TBitmap.Create;
end;

destructor TappliLSE.destroy;
begin
	fBitmap.Free;
  inherited;
end;

procedure TappliLSE.SetExe(const Value: string);
begin
  fexe := Value;
  GetAssociatedBmp(PChar(Value),fbitMap);
end;

procedure TappliLSE.GetAssociatedBmp(const Ext: PChar; var BM : Tbitmap);
var Info: TSHFileInfo;
		HH : TICON;
begin
  HH := TIcon.create;
  SHGetFileInfo(Ext, FILE_ATTRIBUTE_NORMAL, Info, SizeOf(Info), SHGFI_ICON);
  HH.Handle := Info.hIcon;
  TIconToTBitmap (HH,BM);
  HH.Free;
end;

Procedure TappliLSE.TIconToTBitmap(FIcon: TIcon ; var FBitmap : TBitmap);
begin
  with FBitmap, Canvas do
  if Assigned(FIcon) then
  begin
    Width:=FIcon.Width;
    Height:=FIcon.Height;
    Brush.Color:=clNone ;
    FillRect(rect(0,0,Width,Height));
    Draw(0,0,FIcon);
  end else
  begin
    Width:=GetSystemMetrics(SM_CXICON);
    Height:=GetSystemMetrics(SM_CYICON);
    Brush.Color:=clNone ;
    FillRect(rect(0,0,Width,Height));
  end;
end;


{ TlistAppliLse }

function TlistAppliLse.Add(AObject: TappliLSE): Integer;
begin
	result := inherited add(Aobject);
end;

destructor TlistAppliLse.destroy;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := count -1 downto 0 do
    begin
      if TappliLSE(Items [Indice])<> nil then
      begin
      	TappliLSE(Items [Indice]).free;
      	Items [Indice]:= nil;;
      end;
    end;
  end;
  inherited;
end;

function TlistAppliLse.GetItems(Indice: integer): TappliLSE;
begin
  result := TappliLSE (Inherited Items[Indice]);
end;

procedure TlistAppliLse.LanceAppli(Base,user,password : string; Item: Integer);
var TheAppli : TappliLSE;
		Command : string;
    rep,LastRep : string;
begin
	TheAppli := Items[Item];
  Command := TheAppli.fexe;
  Rep := ExtractFilePath(Command);
  if TheAppli.AccepteParam then
  begin
		Command := Command + ' /USER="' + User + '"'+ ' /PASSWORD=' + password + ' /DOSSIER=' + Base;
  	FileExec(Command ,False,False);
  end else
  begin
		Lastrep := GetCurrentDir;
		SetCurrentDir (rep);
	  FileExec(Command ,False,False);
		SetCurrentDir (lastrep);
  end;
  //
end;

procedure TlistAppliLse.SetItems(Indice: integer; const Value: TappliLSE);
begin
  Inherited Items[Indice]:= Value;
end;

end.
