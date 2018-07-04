unit FichePrinc;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, GraphicEx, ExtCtrls, Hctrls, HSysMenu, StdCtrls, HPanel, Mask,USplashLanceur,
  HTB97,CBPPath,Uapplications, Buttons,IniFiles,LicUtil, uPortailCtrls,
  TntStdCtrls, TntExtCtrls;

const INIFILE = 'CEGIDPGI.INI';

type
  TFFichePrinc = class(TForm)
    IMGFOND: THImage;
    HMTRAD: THSystemMenu;
    IMGLSE: THImage;
    TB1: TToolbarButton97;
    TB4: TToolbarButton97;
    TB7: TToolbarButton97;
    TB10: TToolbarButton97;
    TB13: TToolbarButton97;
    TB16: TToolbarButton97;
    TA1: TToolbarButton97;
    TA4: TToolbarButton97;
    TA7: TToolbarButton97;
    TA10: TToolbarButton97;
    TA13: TToolbarButton97;
    TA16: TToolbarButton97;
    TA2: TToolbarButton97;
    TB2: TToolbarButton97;
    TA3: TToolbarButton97;
    TB3: TToolbarButton97;
    TA5: TToolbarButton97;
    TB5: TToolbarButton97;
    TA8: TToolbarButton97;
    TB8: TToolbarButton97;
    TA11: TToolbarButton97;
    TB11: TToolbarButton97;
    TA14: TToolbarButton97;
    TB14: TToolbarButton97;
    TA17: TToolbarButton97;
    TB17: TToolbarButton97;
    TA6: TToolbarButton97;
    TB6: TToolbarButton97;
    TA9: TToolbarButton97;
    TB9: TToolbarButton97;
    TA12: TToolbarButton97;
    TB12: TToolbarButton97;
    TA15: TToolbarButton97;
    TB15: TToolbarButton97;
    TA18: TToolbarButton97;
    TB18: TToolbarButton97;
    HLabel1: THLabel;
    CBSOCIETE: THValComboBox;
    HLabel2: THLabel;
    USER: THCritMaskEdit;
    HLabel3: THLabel;
    PASSWD: THCritMaskEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TB1DblClick(Sender: TObject);
  private
    { Déclarations privées }
    XX : TFsplash;
    PGIiniFile : string;
    ListeAppli : TlistAppliLse;
    TheConnexion : TConnection;
    procedure AffecteBouttons;
    procedure AssocieBouton(indice: Integer; UneAppli: TappliLSE);
    procedure GetSocietes;
  public
    { Déclarations publiques }
  end;

var
  FFichePrinc: TFFichePrinc;

implementation

uses TntStdCtrls;
{$R *.dfm}

function WindowsDirectory: String;
const
  dwLength: DWORD = 255;
var
  WindowsDir: PChar;
begin
  GetMem(WindowsDir, dwLength);
  GetWindowsDirectory(WindowsDir, dwLength);
  Result := String(WindowsDir);
  FreeMem(WindowsDir, dwLength);
end;

function FindPGIIniFile : string;
  var Localrepert : string;
  		Iexist : integer;
begin
	Iexist := -1;
  // recherche sur repertoire courant
  TRY
  	Result := '';
    Localrepert := ExtractFilePath (Application.ExeName); // repertoire de l'application
    Iexist := FileOpen(IncludeTrailingPathDelimiter (Localrepert)+INIFILE,fmOpenRead );
    if Iexist <0 then
    begin
      Localrepert := TCbpPath.GetCegidUserRoamingAppData; // repertoire Utilisateur
      Iexist := FileOpen(IncludeTrailingPathDelimiter(Localrepert)+INIFILE,fmOpenRead );
      if Iexist <0 then
      begin
        Localrepert := TCBPPath.GetCegidData; // Data de All Users
        Iexist := FileOpen(IncludeTrailingPathDelimiter(Localrepert)+INIFILE,fmOpenRead);
        if Iexist < 0 then
        begin
           LocalRepert := WindowsDirectory;         // Emplacement de windows
           Iexist := FileOpen(IncludeTrailingPathDelimiter(Localrepert)+INIFILE,fmOpenRead);
           if Iexist >= 0 then
           begin
              Result := Localrepert;
           end;
        end else
        begin
          Result := Localrepert;
        end;
      end else
      begin
        Result := Localrepert;
      end;
    end else
    begin
      Result := Localrepert;
    end;
  FINALLY
   if Iexist >= 0 then FileClose(Iexist);
  END;
end;

procedure TFFichePrinc.FormCreate(Sender: TObject);
begin
  TheConnexion := TConnection.Create;
  XX := OpenSplash;
	//
  PGIiniFile := FindPGIIniFile;
  if PGIiniFile = '' then
  begin
    MessageDlg('Acune application installée sur ce poste', mtInformation,[mbOk], 0);
    PostMessage(Application.Handle,WM_CLOSE,0,0);
  end;
  //
  ListeAppli := TlistAppliLse.create;
  ConstitueListeAppli( ListeAppli);
  if ListeAppli.Count = 0 then
  begin
    MessageDlg('Acune application installée sur ce poste', mtInformation,[mbOk], 0);
    PostMessage(Application.Handle,WM_CLOSE,0,0);
  end;
  AffecteBouttons;
  GetSocietes;
end;

procedure TFFichePrinc.FormShow(Sender: TObject);
begin
  CloseSplash(XX);
  //

end;

procedure TFFichePrinc.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  TheConnexion.LastSoc := CBSOCIETE.ItemIndex;
  TheConnexion.LastUser := USER.Text;
  SetLastSocSelected (TheConnexion);
  TheConnexion.free;
	CloseSplash(XX);
end;

procedure  TFFichePrinc.AssocieBouton(indice : Integer ;UneAppli : TappliLSE);
var TB : TToolbarButton97;
begin
  TB := TToolbarButton97 (Self.FindComponent('TB'+InttoStr(Indice)));
  if TB <> nil then
  begin
    TB.Visible := True;
    TB.Caption := UneAppli.Titre;
    TB.Glyph.Transparent := True;
    TB.Glyph.TransparentColor  := clNone;
    TB.Update;
    TB.Refresh;
  end;
  TB := TToolbarButton97 (Self.FindComponent('TA'+InttoStr(Indice)));
  if TB <> nil then
  begin
    TB.Visible := True;
    TB.Glyph := UneAppli.BITMap;
    TB.Glyph.Transparent := True;
    TB.Glyph.TransparentColor  := clNone;
    TB.Update;
    TB.Refresh;
  end;
end;


procedure TFFichePrinc.AffecteBouttons;
var  i: integer;
begin
	for i := 0 to ListeAppli.Count -1 do
  begin
    AssocieBouton(i+1,ListeAppli.Items [i]);
  end;
end;

procedure TFFichePrinc.TB1DblClick(Sender: TObject);
var Base,UserName,Password : string;
begin
  base := CBSOCIETE.Text;
  UserName := USER.Text;
  Password := CryptageSt(PASSWD.Text);
  ListeAppli.lanceAppli(base,UserName,Password, TToolbarButton97(Sender).Tag);
end;

procedure TFFichePrinc.GetSocietes;
var FF : TIniFile;
		TT : TStringList;
    i : Integer;
begin
  TT := TStringList.Create;
 	FF := TIniFile.Create(IncludeTrailingPathDelimiter(PGIiniFile)+INIFILE);
  FF.ReadSections(TT);
  for I := 0 to TT.Count -1 do
  begin
    if (UpperCase(TT.strings[i]) <> 'REFERENCE') and (UpperCase(TT.strings[i]) <> '###RESERVED###') then
    begin
    	CBSOCIETE.AddItem(TT.Strings[I],nil )
    end;
  end;
  FF.Free;
  TT.free;
  GetLastSocSelected (TheConnexion);
  CBSOCIETE.ItemIndex := TheConnexion.LastSoc;
  USER.Text := TheConnexion.LastUser;
end;

end.
