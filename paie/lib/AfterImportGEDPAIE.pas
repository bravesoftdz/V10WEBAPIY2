unit AfterImportGEDPAIE;

interface

uses
  Forms,
  UGedFiles,
  HEnt1,
//  PNewDocument,         //ShowNewDocument
  utofPGNewDocument,
  UtomSalaries          //Salarie_GestionBoutonGed
;

 Procedure MyAfterImportPaie (Sender: TObject; FileID: String; var Cancel: Boolean) ;


implementation

 Procedure  MyAfterImportPaie (Sender: TObject; FileID: String; var Cancel: Boolean) ;
// Evenement après scannérisation : retourne le FileID du fichier scanné
var
  ParGed                : TParamGedDoc;
  stInfos               : String;
  stObjet               : string;
begin
  if FileId = '' then
    exit;  // ---- Insertion classique dans la GED avec boite de dialogue ----
  if Sender is TForm then
     begin
     if TForm(Sender).name = 'SALARIE' then
        begin
        stInfos         := Salarie_MyAfterImport(Sender);
        if stInfos <> '' then
          stObjet       := 'SAL';
        end;
     end;
  // Propose le rangement dans la GED
  ParGed.SDocGUID       := '';

  ParGed.NoDossier      := V_PGI.NoDossier;
  ParGed.CodeGed        := '';
  // FileId est le n° de fichier obtenu par la GED suite à l'insertion
  ParGed.SFileGUID      := FileId;
  // Description par défaut du document à archiver...
  if Sender is TForm then
    ParGed.DefName      := TForm(Sender).Caption;
  ParGed.Objet          := stObjet;
  ParGed.Infos          := stInfos;
  if stObjet = '' then
    ParGed.ModifLien    := True
  else
    ParGed.ModifLien    := False;
  Application.BringToFront;
  if ShowNewDocument(ParGed) = '##NONE##' then
     // Fichier refusé, suppression dans la GED
     V_GedFiles.Erase(FileId)
  else
     begin
       if (TForm(Sender).name = 'SALARIE') and (stInfos <> '') then
        Salarie_GestionBoutonGed(Sender);
     end;
end;

end.
