unit AfterImportGEDGRC;

interface

uses
  Forms,
  UGedFiles,
  HEnt1,
  RTNewDocument,
  utomtiers,
  utomactionchaine,
  utomaction,
  utomperspective
;

 Procedure MyAfterImportGRC (Sender: TObject; FileID: String; var Cancel: Boolean) ;


implementation

 Procedure  MyAfterImportGRC (Sender: TObject; FileID: String; var Cancel: Boolean) ;
// Evenement après scannérisation : retourne le FileID du fichier scanné
var ParGed : TParamGedDoc;
    stInfos,stObjet : string;
begin
  if FileId = '' then exit;  // ---- Insertion classique dans la GED avec boite de dialogue ----
  if Sender is TForm then
     begin
     if TForm(Sender).name = 'GCTIERS' then
        begin
        stInfos := Tiers_MyAfterImport(Sender);
        if stInfos <> '' then stObjet := 'TIE';
        end;
     if TForm(Sender).name = 'RTACTIONSCHAINE' then
        begin
        stInfos := Chainage_MyAfterImport(Sender);
        if stInfos <> '' then stObjet := 'CHA';
        end;
     if (TForm(Sender).name = 'RTACTIONS') or (TForm(Sender).name = 'RTCHAINAGES') then
        begin
        stInfos := Action_MyAfterImport(Sender);
        if stInfos <> '' then stObjet := 'ACT';
        end;
     if TForm(Sender).name = 'RTPERSPECTIVES' then
        begin
        stInfos := Perspective_MyAfterImport(Sender);
        if stInfos <> '' then stObjet := 'PRO';
        end;
     end;
  // Propose le rangement dans la GED
  ParGed.SDocGUID := '';

  ParGed.NoDossier := V_PGI.NoDossier;
  ParGed.CodeGed := '';
  // FileId est le n° de fichier obtenu par la GED suite à l'insertion
  ParGed.SFileGUID := FileId;
  // Description par défaut du document à archiver...
  if Sender is TForm then ParGed.DefName := TForm(Sender).Caption;
  ParGed.Objet := stObjet;
  ParGed.Infos := stInfos;
  if stObjet = '' then ParGed.ModifLien := True else ParGed.ModifLien := False;
  Application.BringToFront;
  if ShowNewDocument(ParGed)='##NONE##' then
     // Fichier refusé, suppression dans la GED
     V_GedFiles.Erase(FileId)
  else
     begin
     if ((TForm(Sender).name = 'RTACTIONS') or (TForm(Sender).name = 'RTCHAINAGES')) and (stInfos <> '') then Action_GestionBoutonGed(Sender)
     else if (TForm(Sender).name = 'RTACTIONSCHAINE') and (stInfos <> '') then Chainage_GestionBoutonGed(Sender)
     else if (TForm(Sender).name = 'RTPERSPECTIVES') and (stInfos <> '') then Perspective_GestionBoutonGed(Sender)
     else if (TForm(Sender).name = 'GCTIERS') and (stInfos <> '') then Tiers_GestionBoutonGed(Sender);
     end;
end;

end.
 