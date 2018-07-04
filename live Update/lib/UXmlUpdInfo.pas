unit UXmlUpdInfo;

interface
uses XmlDoc,XMLIntf,Windows,UFolders,classes;

const
	XMLNAME = 'UPDPRODUCT';
type
	TXmlUpdInfo = class
  	private
      fDomDoc : TXmlDocument;
      fParentNode : IXmlnode;
      fOkData : boolean;
      fXmlFileFolder : widestring;
    public
    	constructor create; overload;
    	constructor create (From : TComponent) ; overload;
      destructor destroy; override;
      procedure Addproduct (Nom, Code, versionInst, versionToUpd , From: string);
      function IsEmpty : boolean;
      function GetNbProducts : integer;
      procedure GetProduct (Indice : integer; var Nom ,Code, from : string); 
      procedure NewXml;
      procedure loadXml;
      procedure SaveXml;
      procedure DeleteXml;
  end;


implementation

uses SysUtils;

{ TXmlUpdInfo }

procedure TXmlUpdInfo.Addproduct(Nom, Code, versionInst, versionToUpd, From : string);
var inodeFille : IXmlnode;
begin
  iNodeFille := fparentnode.AddChild('PRODUCT');
  iNodeFille.Attributes ['Code'] := Code;
  iNodeFille.Attributes ['Name'] := Nom;
  iNodeFille.Attributes ['Installed'] := versionInst;
  iNodeFille.Attributes ['ToInstal'] := versionToUpd;
  iNodeFille.Attributes ['From'] := From;
  fOkData := true;
end;

constructor TXmlUpdInfo.create;
begin
  fOkDatA := false;
	fDomDOc := TXmlDocument.create(nil);
  fXmlFileFolder := IncludeTrailingBackslash (GetSpecialFolders (CSIDL_COMMON_APPDATA))+'LseUpdate';
end;

constructor TXmlUpdInfo.create(From: TComponent);
begin
  fOkDatA := false;
	fDomDOc := TXmlDocument.create(From);
  fXmlFileFolder := IncludeTrailingBackslash (GetSpecialFolders (CSIDL_COMMON_APPDATA))+'LseUpdate';
end;

procedure TXmlUpdInfo.DeleteXml;
var NomXml : string;
begin
	NomXml := IncludeTrailingBackslash (fXmlFileFolder)+XMLNAME+'.xml';
  DeleteFile(NomXml);
end;

destructor TXmlUpdInfo.destroy;
begin
  fdomdoc.active := false;
  inherited;
end;

function TXmlUpdInfo.GetNbProducts: integer;
var II : IXMLNode;
begin
  result := 0;
  result :=  fDomDoc.DocumentElement.ChildNodes.count;
  sleep(100)
end;

procedure TXmlUpdInfo.GetProduct(Indice: integer; var Nom, Code, from : string);
var TheNode : IXmlnode;
begin
	Nom := '';
  Code := '';
  From := '';
  theNode := fDomDoc.DocumentElement.ChildNodes.Get (Indice);
  if TheNode = nil then exit;
  Nom := TheNode.GetAttributeNS('Name','');
  Code := TheNode.GetAttributeNS('Code','');
  From := TheNode.GetAttributeNS('From','');
end;

function TXmlUpdInfo.IsEmpty: boolean;
begin
	result := FDomDoc.IsEmptyDoc;
end;

procedure TXmlUpdInfo.loadXml;
var NomXml : string;
begin
	NomXml := IncludeTrailingBackslash (fXmlFileFolder)+XMLNAME+'.xml';
  if not FileExists(NomXml) then exit;
  fDomDoc.LoadFromFile(NomXml);
  fdomdoc.active := true;
end;

procedure TXmlUpdInfo.NewXml;
begin
  fdomdoc.active := true;
  fdomdoc.Version := '1.0';
  fDomDoc.Encoding := 'UTF-8';
  fDomdoc.Options := [doNodeAutoIndent];
  fParentNode := fdomDoc.CreateElement('PRODUCTS','');
  fDomDoc.ChildNodes.Add(fParentNode);
end;

procedure TXmlUpdInfo.SaveXml;
var NomXml : string;
begin
	if not FokData then exit;
	NomXml := IncludeTrailingBackslash(fXmlFileFolder)+XMLNAME+'.xml';
	fDomDoc.SaveToFile(NomXml)
end;

end.
