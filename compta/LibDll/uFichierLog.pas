{ Unité : Génération d'un fichier de log
  Deux possibilités :
   1/ On stock tous les logs et on écrit le fichier en fin de traitement avec réinitialisation du fichier
      => Utiliser Add() puis GenereFichier
   2/ On écrit au fur et à mesure Ecrit
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 8.01.001.006   01/03/07    JP     Création de l'unité
--------------------------------------------------------------------------------------}
unit uFichierLog;

interface

uses
  Classes;

type
  TObjFichierLog = class
  private
    ListeMsg : TStringList;
    FFichier : string;
    FReInit  : Boolean;
  public
    destructor  Destroy; override;
    {NomAppli : Nom du fichier; AvecReInit à True, on vide le fichier dans GenereFichier}
    constructor Create(const NomAppli : string; AvecReInit : Boolean = False);
    procedure   Add   (const stLog : string);
    procedure   Ecrit (const stLog : string);
    procedure   GenereFichier;
  end;

var
  ObjLog : TObjFichierLog;

implementation

uses
  SysUtils, Forms;

{ TObjFichierLog }

{---------------------------------------------------------------------------------------}
constructor TObjFichierLog.Create(const NomAppli : string; AvecReInit : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin
  try
    ListeMsg := TStringList.Create;
    if Trim(NomAppli) <> '' then
      FFichier := ExtractFileDir(Application.ExeName) + '\LOG_' + NomAppli + '.TXT'
    else begin
      DecodeDate(Date, a, m, j);
      FFichier := ExtractFileDir(Application.ExeName) + '\LOG_' +
                  IntToStr(a) + '_' +
                  IntToStr(m) + '_' +
                  IntToStr(j) + '.TXT';
    end;
    FReInit  := AvecReInit or not FileExists(FFichier);
  except
    on E : Exception do 
      raise Exception.Create('CreateLog' + #13#13 + E.Message);
  end;
end;

{---------------------------------------------------------------------------------------}
destructor TObjFichierLog.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ListeMsg) then FreeAndNil(ListeMsg);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFichierLog.Add(const stLog : string);
{---------------------------------------------------------------------------------------}
begin
  ListeMsg.Add(stLog);
end;

{---------------------------------------------------------------------------------------}
procedure TObjFichierLog.Ecrit(const stLog : string);
{---------------------------------------------------------------------------------------}
var
  TF : TextFile;
begin
  try
    AssignFile(TF, FFichier);
    try
      if not FileExists(FFichier) then Rewrite(TF)
                                  else Append(TF);
      Writeln(TF, stLog);
      Flush(TF);
    finally
      CloseFile(TF);
    end;
  except
    on E : Exception do
      raise Exception.Create('Log impossible : ' + FFichier + #13#13 + E.Message);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFichierLog.GenereFichier;
{---------------------------------------------------------------------------------------}
var
  TF : TextFile;
  n  : Integer;
begin
  try
    AssignFile(TF, FFichier);
    try
      if FReInit then Rewrite(TF)
                 else Append(TF);
      if not FReInit then begin
        Writeln(TF, '');
        Writeln(TF, '-----------------------------------------------------------------------');
        Writeln(TF, '           DATE : ' + DateTimeToStr(Now));
        Writeln(TF, '-----------------------------------------------------------------------');
        Writeln(TF, '');
      end;

      for n := 0 to ListeMsg.Count - 1 do
        Writeln(TF, ListeMsg[n]);

      Flush(TF);
      ListeMsg.Clear;
    finally
      CloseFile(TF);
    end;
  except
    on E : Exception do
      raise Exception.Create('Log impossible : ' + FFichier + #13#13 + E.Message);
  end;
end;

end.
