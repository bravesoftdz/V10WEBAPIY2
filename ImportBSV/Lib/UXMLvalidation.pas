unit UXMLvalidation;

interface

uses
  Classes, SysUtils, msxmldom,
  MSXML2_TLB, { Il faut importer l'activeX MSXML DOM2 4.0 pour obtenir l'unité MSXML2_TLB.pas }
  xmldom,
  XMLDoc, XMLIntf; { Pour une utilisation de TXMLDocument }

type
  EXMLDOMValidationError = class(Exception)
  private
    FSrcText: DOMString;
    FURL: DOMString;
    FReason: DOMString;
    FErrorCode: Integer;
    FLine: Integer;
    FLinePos: Integer;
    FFilePos: Integer;
  public
    constructor Create(const ValidationError: IXMLDOMParseError);
    property ErrorCode: Integer read FErrorCode;
    property URL: DOMString read FURL;
    property Reason: DOMString read FReason;
    property SrcText: DOMString read FSrcText;
    property Line: Integer read FLine;
    property LinePos: Integer read FLinePos;
    property FilePos: Integer read FFilePos;
  end;


procedure ValidateXMLFileName(XMLFilename,  XSDFilename: TFileName; NameSpaceSchema: string);
procedure ValidateXMLText(XMLText: string; XSDFilename: TFileName; NameSpaceSchema: string);
procedure ValidateXMLDoc(XMLDoc, XSDDoc: IXMLDOMDocument2;   NameSpaceSchema: string);

// Génère une exception EXMLDOMValidationError si errorcode<>0
procedure CheckValidationError(ValidationError: IXMLDOMParseError);

// Permet de retrouver l'interface COM de MSXML à partir d'un IXMLDocument (Borland)
// Attention: le IXMLDocument doit avoir dans ce cas MSXML comme Vendor !
function GetXMLDOMDocument2(XMLDoc: IXMLDocument): IXMLDOMDocument2;

resourcestring
  RS_EXMLDOM_FILE_VALIDATION_ERROR = 'Document XML non conforme au schéma.'#13#10#13#10+
                                     'Raison:'#13#10'%s'#13#10+
                                     'Ligne: %d'#13#10+
                                     'Position: %d'#13#10+
                                     'Position dans le fichier: %d'#13#10+
                                     'URL: %s'#13#10+
                                     'XML: '#13#10'%s';

  RS_EXMLDOM_NOFILE_VALIDATION_ERROR = 'Document XML non conforme au schéma.'#13#10#13#10+
                                       'Raison:'#13#10'%s'#13#10;

implementation

{ EXMLDOMValidationError }

constructor EXMLDOMValidationError.Create(
  const ValidationError: IXMLDOMParseError);
begin
  FSrcText := ValidationError.srcText;
  FURL := ValidationError.url;
  FReason := ValidationError.reason;
  FErrorCode := ValidationError.errorCode;
  FLine := ValidationError.line;
  FLinePos := ValidationError.linepos;
  FFilePos := ValidationError.filepos;

  if FLine>0 then
    inherited CreateResFmt(@RS_EXMLDOM_FILE_VALIDATION_ERROR, [FReason,
                                                               FLine,
                                                               FLinePos,
                                                               FFilePos,
                                                               FURL,
                                                               FSrcText])
  else
    inherited CreateResFmt(@RS_EXMLDOM_FILE_VALIDATION_ERROR, [FReason]);
end;

function GetXMLDOMDocument2(XMLDoc: IXMLDocument): IXMLDOMDocument2;
begin
  XMLDoc.Active := True;
  Result := ((XMLDoc.DOMDocument as IXMLDOMNodeRef).GetXMLDOMNode as IXMLDOMDocument2);
end;

procedure CheckValidationError(ValidationError: IXMLDOMParseError);
begin
  if ValidationError.errorCode<>0 then
    raise EXMLDOMValidationError.Create(ValidationError);
end;

procedure ValidateXMLDoc(XMLDoc, XSDDoc: IXMLDOMDocument2;
  NameSpaceSchema: string);
var
  vSchemaCollection: IXMLDOMSchemaCollection2;
begin
  vSchemaCollection := CoXMLSchemaCache40.Create;
  vSchemaCollection.add(NameSpaceSchema, XSDDoc);
  XMLDoc.schemas := vSchemaCollection;
  CheckValidationError(XMLDoc.validate);
end;

procedure ValidateXMLFilename(XMLFilename,
  XSDFilename: TFileName; NameSpaceSchema: string);
var
  vXML,
  vSchema: IXMLDOMDocument2;
  vSchemaCollection: IXMLDOMSchemaCollection2;
begin
  vSchema := CoDOMDocument40.Create;
  vSchema.async := False;
  vSchema.load(XSDFilename);

  vSchemaCollection := CoXMLSchemaCache40.Create;
  vSchemaCollection.add(NameSpaceSchema, vSchema);

  vXML := CoDOMDocument40.Create;
  vXML.async := False;
  vXML.validateOnParse := True;
  vXML.schemas := vSchemaCollection;
  vXML.load(XMLFilename);

  CheckValidationError(vXML.parseError);
end;

procedure ValidateXMLText(XMLText: string;
  XSDFilename: TFileName; NameSpaceSchema: string);
var
  vXML,
  vSchema: IXMLDOMDocument2;
  vSchemaCollection: IXMLDOMSchemaCollection2;
begin
  vSchema := CoDOMDocument40.Create;
  vSchema.async := False;
  vSchema.load(XSDFilename);

  vSchemaCollection := CoXMLSchemaCache40.Create;
  vSchemaCollection.add(NameSpaceSchema, vSchema);

  vXML := CoDOMDocument40.Create;
  vXML.async := False;
  vXML.validateOnParse := True;
  vXML.schemas := vSchemaCollection;
  vXML.loadXML(XMLText);

  CheckValidationError(vXML.parseError);
end;

 
end.
