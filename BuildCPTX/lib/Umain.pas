unit Umain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, HTB97, ExtCtrls, StdCtrls, TntStdCtrls, Hctrls,Aglinit,hmsgbox,HEnt1,GalPatience,ShellApi,UCOntroleUAC;

type
  TForm1 = class(TForm)
    CreBuild: TToolbarButton97;
    HLabel1: THLabel;
    procedure CreBuildClick(Sender: TObject);
  private
    { Déclarations privées }
    function SauveAncienBuild : boolean;
    procedure VideRepertoire (Repert : string);
    procedure CopieRepertoire (FromRepert,ToRepert : string);
    function NouveauBuild: boolean;
  public
    { Déclarations publiques }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}




function TForm1.NouveauBuild : boolean;
begin
  TRY
    Result := True;
    VideRepertoire('C:\CptxBuilder\CPTX');
    VideRepertoire('C:\CptxBuilder\DllDelphi\Bin');
    VideRepertoire('C:\CptxBuilder\DllDelphi\Dcu');
    VideRepertoire('C:\CptxBuilder\DllDelphi\Syn');
    VideRepertoire('C:\CptxBuilder\DllDelphi\Lib');
    VideRepertoire('C:\CptxBuilder\DllDelphi\Syn');
    VideRepertoire('C:\CptxBuilder\DemandeSocREF\BOBSBTP');
    VideRepertoire('C:\CptxBuilder\DemandeSocREF\SocRefBTP');
    CopieRepertoire('C:\CPTXCONSTRUCT\trunk\Organisation de base (CptxBuilder)\DllDelphi\Lib','C:\CptxBuilder\DllDelphi\Lib');
  EXCEPT
    Result := false;
  end;

end;

function TForm1.SauveAncienBuild : boolean;
var Dir : string;
    CptxFile,FromFile,DestFile : string;
    Info : TSearchRec;
begin
  Result := false;
  DIr := IncludeTrailingPathDelimiter('C:\CptxBuilder\CPTX');
  CptxFile := '';
  If FindFirst(DIr+'*.CPTX',faAnyFile,Info)=0 Then
  Begin
    Repeat
      If (Info.Attr And faDirectory)=0 then
      begin
        CptxFile := Info.findData.cFileName;
        break;
      end;
      { Il faut ensuite rechercher l'entrée suivante }
    Until FindNext(Info)<>0;
    FindClose(Info);
  End;
  if CptxFile <> '' then
  begin
    FromFile := 'C:\CptxBuilder';
    Destfile := IncludeTrailingPathDelimiter( 'C:\CPTXCONSTRUCT\tags')+copy(CptxFile,1,Length(CptxFile)-5) ;
    if not DirectoryExists(DestFile) then
    begin
      CopieRepertoire (FromFile,DestFile);
      Result := True;
    end;
  end;
end;

procedure TForm1.CreBuildClick(Sender: TObject);
var XX : TFPatience;
    okok : Boolean;
begin
  if Pgiask ('ATTENTION : Cette action va générer un nouveau build pour un CPTX.#13#10 Confirmez-vous l''action ?') <> mryes then exit;
  XX := FenetrePatience ('Constitution du nouveau Build CPTX');
  TRY
    XX.lCreation.Caption := 'Sauvegarde de l''ancien BUILD';
    XX.invalidate;
    XX.StartK2000 ;
    Application.ProcessMessages;
    okok := SauveAncienBuild;
    if okok then
    begin
      XX.lCreation.Caption := 'Constitution du nouveau BUILD';
      XX.Invalidate;
      Application.ProcessMessages;
      if NouveauBuild then
      begin
        PgiInfo('Nouveau build créé avec succès');
      end;
    end else
    begin
      PgiError ('Traitement aborté.#13#10 le Build précédent ne peut être sauvegardé.#13#10 Merci de contrôler');
    end;
  FINALLY
    XX.free;
  END;
end;

procedure TForm1.CopieRepertoire (FromRepert,ToRepert : string);
begin
  ShellExecAndWait('c:\windows\system32\xcopy', '"'+FromRepert+'"  "'+ToRepert+'" /E /C /I /Y',SW_HIDE);
end;


procedure TForm1.VideRepertoire(Repert: string);
var Info : TSearchRec;
    FFile : string;
begin
  If FindFirst(IncludeTrailingPathDelimiter (Repert)+'*.*',faAnyFile,Info)=0 Then
  Begin
    Repeat
      If (Info.Attr And faDirectory)=0 then
      begin
        DeleteFile(IncludeTrailingPathDelimiter (Repert)+Info.findData.cFileName);
      end else
      begin
        if (Info.Name  <> '.') and (Info.Name <> '..') then
        begin
          VideRepertoire (IncludeTrailingPathDelimiter (Repert)+Info.findData.cFileName);
          RemoveDir(IncludeTrailingPathDelimiter (Repert)+Info.findData.cFileName);
        end;
      end;
      { Il faut ensuite rechercher l'entrée suivante }
    Until FindNext(Info)<>0;
    FindClose(Info);
  End;

end;

end.



