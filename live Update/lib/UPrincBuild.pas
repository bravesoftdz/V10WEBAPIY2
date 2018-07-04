unit UPrincBuild;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, ExtCtrls, StdCtrls, Mask, FileCtrl,UbrowseRepert, Grids,
  ComCtrls,UConstantes,UEncryptage,Ulog;

type
  TMainForm = class(TForm)
    Panel1: TPanel;
    BtCache: TSpeedButton;
    Bvalide: TSpeedButton;
    Label1: TLabel;
    Repository: TEdit;
    BFindEmplacement: TSpeedButton;
    Panel2: TPanel;
    LanceTrait: TSpeedButton;
    RPT: TRichEdit;
    procedure BtCacheClick(Sender: TObject);
    procedure BFindEmplacementClick(Sender: TObject);
    procedure LanceTraitClick(Sender: TObject);
  private
    { Déclarations privées }
    procedure BuildRepository;
    procedure BuildRepositoryDetail (SubDirectory : string);
    procedure TraiteFile(NameFile, SubDirectory: string);

  public
    { Déclarations publiques }
  end;

var
  MainForm: TMainForm;

implementation

uses Math,UInprogress;

{$R *.dfm}

procedure TMainForm.BtCacheClick(Sender: TObject);
begin
  close;
end;

procedure TMainForm.BFindEmplacementClick(Sender: TObject);
var TheRepert : string;
    DLG1 : TopenDialog;
begin
  TheRepert := Repository.text;
  if SelectDirectory ('Séléctionner un dépot',Repository.text,Therepert) then Repository.text := TheRepert;
end;

procedure TMainForm.LanceTraitClick(Sender: TObject);
var XX : TFInprogress;
begin
  if Application.MessageBox('Vous allez constituer le fichier d''information du dépot.Confirmez-vous ?','Demande de Confirmation',MB_OKCANCEL or MB_ICONEXCLAMATION) <> IDOK then exit;
  XX := TFInprogress.create(Application);
  XX.Show;
  XX.Refresh;
  DeleteFile(IncludeTrailingBackslash(Repository.text)+CONFLSEFILE);
  DeleteFile(IncludeTrailingBackslash(Repository.text)+CONFCEGIDFILE);
  BuildRepository;
  XX.free;
end;

procedure TMainForm.TraiteFile(NameFile,SubDirectory : string);
var NameConfig : string;
    InfoFile : string;
begin
  if Pos(SubDirectory,'PROGRAMDATA;BOBS2009;LSEAPP')>0 then NameConfig := CONFLSEFILE
                                                       else NameConfig := CONFCEGIDFILE;
  InfoFile := NameFile+SEPARATOR+SUBDIRECTORY+SEPARATOR+MD5(IncludeTrailingBackslash(Repository.text)+IncludeTrailingBackslash(SubDirectory)+NameFIle);
  ecritLog (IncludeTrailingBackslash(Repository.text),InfoFile,NameConfig);
end;

procedure TMainForm.BuildRepositoryDetail (SubDirectory : string);
var Info   : TSearchRec;
begin
  If FindFirst(IncludeTrailingBackslash(Repository.text)+IncludeTrailingBackslash(SubDirectory)+'*.*',faAnyFile,Info)=0 Then
  Begin
    Repeat
      If (info.Name<>'.')And(info.Name<>'..') then
      begin
        If ((Info.Attr and faDirectory)<>faDirectory) then
        begin
          TraiteFile(Info.Name,SubDirectory);
        end;
      end;
    { Il faut ensuite rechercher l'entrée suivante }
    Until FindNext(Info)<>0;
  end;
end;

procedure TMainForm.BuildRepository;
var TL : TStringList;
    Info   : TSearchRec;
    II : integer;
begin
  TL := TStringList.Create;
  TRY
    { Recherche de la première entrée du répertoire }
    If FindFirst(IncludeTrailingBackslash(Repository.text)+'*.',faAnyFile,Info)=0 Then
    Begin
      Repeat
        If (info.Name<>'.')And(info.Name<>'..') then
        begin
          If ((Info.Attr and faDirectory)=faDirectory) then
          begin
            TL.Add(extractFileName(info.name));
          end;
        end;
      { Il faut ensuite rechercher l'entrée suivante }
      Until FindNext(Info)<>0;
    end;
    FindClose(Info);
    if TL.Count > 0 then
    begin
      for II := 0 to TL.Count -1 do
      begin
        BuildRepositoryDetail (TL.Strings[II]);
      end;
    end;
  FINALLY
    TL.free;
  END;
end;

end.
