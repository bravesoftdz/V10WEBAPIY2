// $$$ 28/07/04 - gestion callbacks d'envoi vers publifi edition
unit LinkPubliFi;

interface

uses
    controls,
    hpdfviewer,
    hctrls;

type
    TLinkPubliFi = class (TObject)
    public
          function  BeforePubliFiPrint (const PDFPreview:TWinControl; var PSFileName:Hstring; var PrinterName:string):boolean;
          procedure AfterPubliFiPrint  (const PDFPreview:TWinControl; PSFilePath:string);
    end;

var
   lkPubliFi   :TLinkPubliFi;

implementation

uses sysutils, forms, hent1, printers, hmsgbox;

function TLinkPubliFi.BeforePubliFiPrint (const PDFPreview:TWinControl; var PSFileName:Hstring; var PrinterName:string):boolean;
begin
     // On construit un nom de fichier exploitable par publifi edition
     PSFileName := 'PGIExpert_' + ChangeFileExt (ExtractFileName (application.ExeName), '') + '_' + PSFileName; // TitreHalley + '_' + PSFileName;

     // Autorisation d'envoyer vers PubliFi Edition uniquement si l'imprimante publifi edition ps existe
     PrinterName := 'PubliFi Edition PS';
     if Printer.Printers.IndexOf (PrinterName) = -1 then
     begin
          PgiInfo ('Envoi vers PubliFi Edition impossible, car le pilote d''imprimante "PubliFi Edition PS" est absent');
          Result := FALSE;
     end
     else
         Result := TRUE;
end;

procedure TLinkPubliFi.AfterPubliFiPrint (const PDFPreview:TWinControl; PSFilePath:string);
begin
     PgiInfo ('Envoi vers PubliFi Edition terminé');
end;

initialization
              lkPubliFi := TLinkPubliFi.Create;

finalization
              lkPubliFi.Free;


end.
