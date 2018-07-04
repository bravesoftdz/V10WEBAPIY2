unit UtilGedRT;

interface

uses HCtrls, SysUtils, UTob, UGedFiles, HMsgBox
     ;


procedure SupprimeDocumentGedRT(sDocId: String);
function  ExtraitDocumentRT(sDocId : String; var sTitre: String): String;

//////////// IMPLEMENTATION ////////////
implementation

uses UtilGed;

procedure SupprimeDocumentGedRT(sDocId: String);
begin
  // pour l'instant pas besoin de tom sur DPDOCUMENT (niveau le plus haut)
  ExecuteSQL('DELETE FROM RTDOCUMENT WHERE RTD_DOCGUID="' + sDocId + '"');

  // suppression dans YDOCUMENTS, YDOCFILES, YFILES, YFILEPARTS :
  SupprimeDocument(sDocId);
end;


function ExtraitDocumentRT(sDocId : String; var sTitre: String): String;

// Retourne le chemin complet du fichier principal du document,
// qui est extrait de la GED dans le rép. temporaire
var TobDoc, TEnreg : TOB;
    sTempFileName : String;
begin
  Result := '';

  // Chargement tob de l'enreg
  TobDoc := TOB.Create('Les documents', Nil, -1);
  // INNER car il y a toujours un document, même s'il n'a pas de fichiers
  TobDoc.LoadDetailFromSQL('SELECT RTDOCUMENT.*, YDOCUMENTS.*, YDOCFILES.*, YFI_FILEGUID, YFI_FILENAME'
   + ' FROM RTDOCUMENT INNER JOIN YDOCUMENTS ON RTD_DOCGUID=YDO_DOCGUID'
                   + ' LEFT JOIN YDOCFILES ON YDO_DOCGUID=YDF_DOCGUID'
                   + ' LEFT JOIN YFILES ON YDF_FILEGUID=YFI_FILEGUID'
   + ' WHERE RTD_DOCGUID="'+sDocId+'" AND YDF_FILEROLE="PAL"');

  if TobDoc.Detail.Count=0 then
    begin
    PGIInfo('Document n° '+sDocId+' non trouvé.', 'Extraction d''un document');
    TobDoc.Free;
    exit;
    end;


  // #### 1 seul enreg, tant qu'on ne fait référence qu'à un fichier PRINCIPAL du document
  TEnreg := TobDoc.Detail[0];

  // #### AGL580 : if VarIsNull(TEnreg.GetValue('YFI_FILEID')) then
  // Attention : sur un type entier, le GetString renvoit '0'
  if TEnreg.GetString('YFI_FILEGUID')='' then
    begin
    PGIInfo('Pas de fichier associé au document n° '+TEnreg.GetValue('YDO_DOCGUID'),
      TEnreg.GetValue('YDO_LIBELLEDOC'));
    TobDoc.Free;
    exit; // SORTIE (pas de document)
    end;

  // #### que mettre en titre ?
  sTitre := 'Document - '+TEnreg.GetValue('YDO_LIBELLEDOC')+' - '+TEnreg.GetValue('YDO_ANNEE');

  // Nom du fichier temporaire
  sTempFileName := V_GedFiles.TempPath + TEnreg.GetValue('YFI_FILENAME');
  // si la tentative échoue avec le nom d'origine
  if FileExists(sTempFileName) then
    // on récupère un nom temporaire quelconque, avec la bonne extension
    sTempFileName := V_GedFiles.TempFileName(ExtractFileExt(TEnreg.GetValue('YFI_FILENAME')));
  // Extraction
  if V_GedFiles.Extract(sTempFileName, TEnreg.GetString('YFI_FILEGUID')) then
    Result := sTempFileName;

  TobDoc.Free;
end;

end.
