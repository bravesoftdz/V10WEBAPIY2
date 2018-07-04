unit UmajCli;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls,UXmlUpdInfo, XPMan, Uregistry,UfileCtrl,UcontroleUAC,ShellApi;

type
  TFmajCli = class(TForm)
    ProgressCopie: TProgressBar;
    Label1: TLabel;
    LProduct: TLabel;
    XPManifest1: TXPManifest;
  private
    { Déclarations privées }

  public
    { Déclarations publiques }
  end;

procedure LanceMajCli;

implementation

{$R *.dfm}

procedure CopyLanceur;
var ServerLocation : string;
		Locallocation : string;
begin
	serverLocation := GetInfoStocke('ServerFolder');
  Locallocation := GetInfoStocke('LocalFolder');
  if not FileExists(IncludeTrailingBackslash(Serverlocation)+'KIT') then
  begin
  	Creationdir (IncludeTrailingBackslash(Serverlocation)+'KIT');
  end;
  if not FileExists(IncludeTrailingBackslash(IncludeTrailingBackslash(Serverlocation)+'KIT')+'UPDATE') then
  begin
  	Creationdir (IncludeTrailingBackslash(IncludeTrailingBackslash(Serverlocation)+'KIT')+'UPDATE');
  end;
  // pour le cas ou ...
  CopieFichier(IncludeTrailingBackslash(Serverlocation)+'KIT\Setup.xml',IncludeTrailingBackslash(Locallocation)+'KIT\Setup.xml');
  CopieFichier( IncludeTrailingBackslash(Serverlocation)+'KIT\SetupCegid.exe',IncludeTrailingBackslash(Locallocation)+'KIT\SetupCegid.exe');
end;

procedure CopyMsi (Code : string; XX : TFMajCli; Indice : integer);
var ServerLocation : string;
		Locallocation : string;
begin
	serverLocation := GetInfoStocke('ServerFolder');
  Locallocation := GetInfoStocke('LocalFolder');
  CopieFichier( IncludeTrailingBackslash(Serverlocation)+'KIT\PRODUCTS\'+Code+'.msi', IncludeTrailingBackslash(locallocation)+'KIT\PRODUCTS\'+Code+'.msi');
  XX.ProgressCopie.Position := Indice;
  XX.Refresh;
  CopieFichier( IncludeTrailingBackslash(Serverlocation)+'KIT\PRODUCTS\'+Code+'.xml', IncludeTrailingBackslash(locallocation)+'KIT\PRODUCTS\'+Code+'.xml');
end;

procedure InstallMsi (FromServer : boolean);
var Locallocation : string;
		Commande : string;
begin
	if fromserver then Locallocation := IncludeTrailingBackslash(GetInfoStocke('ServerFolder'))+'KIT'
  							else Locallocation := IncludeTrailingBackslash(GetInfoStocke('LocalFolder'))+'KIT';
  Commande := IncludeTrailingBackslash(LocalLocation)+'SetupCegid.exe';
  if ShellExecute(application.Handle,'open',PChar(commande),nil,nil,SW_SHOWNORMAL) <= 32 then RaiseLastOSError;
end;

procedure LanceMajCli;
var Nbproducts : integer;
    Code : string;
    Libelle : string;
    From : string;
    XmlDoc : TXmlUpdInfo;
    XX : TFMajCli;
    indice : integer;
    ServerUpd : boolean;
begin
	ServerUpd := false;
  if getInfoStocke('UpdMethod')='0' then ServerUpd := true;
	XmlDoc := TXmlUpdInfo.create (Application);
  XmlDoc.loadXml;
  if XmlDoc.IsEmpty then exit;
  NbProducts := XmlDoc.GetNbProducts;
  if Nbproducts = 0 then exit;
  CopyLanceur;
  if not ServerUpd then
  begin
    XX := TFMAJCli.Create(application);
    TRY
      XX.ProgressCopie.Min := 0;
      XX.ProgressCopie.Max := nbproducts;
      XX.ProgressCopie.StepBy(1);
      if ServerUpd then XX.ProgressCopie.visible := false;
      for indice := 0 to nbproducts-1 do
      begin
        XmlDoc.GetProduct(indice,Libelle,Code,From);
        XX.Lproduct.Caption := Libelle;
        if indice = 0 then XX.Show ;
        XX.Refresh;
        if not Serverupd then CopyMsi (Code,XX,Indice);
        XX.ProgressCopie.Position := Indice+1;
        XX.Refresh;
      end;
    FINALLY
      XX.free;
    END;
  end;
  InstallMsi(ServerUpd);
  xmlDoc.DeleteXml;
  xmlDoc.free;
end;

end.
