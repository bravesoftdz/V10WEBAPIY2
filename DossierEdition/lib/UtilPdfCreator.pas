unit UtilpdfCreator;

interface

uses Windows,
     classes,
     Graphics,
     math;

type

TpdfCreator = class (Tobject)
	private
  	fWinPdfCreat : olevariant;
    fWinExcel : Olevariant;
    fWinWord : Olevariant;
    fOk : boolean;
  public
  	constructor create;
    destructor destroy; override;
    procedure Open (newInst : boolean);
    procedure PrepareWorkBook;
    procedure AddExcelFile (NomFichier : string);
end;

implementation

uses sysutils, dialogs, forms, comobj, ActiveX, hctrls, hent1, controls,hmsgbox,
  Variants,UtilXlsBtp;

{ TpdfCreator }

procedure TpdfCreator.AddExcelFile(NomFichier: string);
var FOpen : boolean;
begin
  OpenExcel (True,fWinExcel,FOpen);
  if VarIsEmpty(fWinExcel) then
  begin
  	V_PGI.IOError := oeUnknown;
  	exit;
  end;
  TRY

  FINALLY
  	if not varIsEmpty(fWinExcel) then fWinExcel.Quit;
    fWinExcel := Unassigned;
  END;

end;

constructor TpdfCreator.create;
begin
	fWinPdfCreat := Unassigned;
  fOk := false;
end;

destructor TpdfCreator.destroy;
begin
	if not VarIsEmpty(fWinPdfCreat) then fWinPdfCreat := UnAssigned;
  inherited;
end;


procedure TpdfCreator.Open(newInst: boolean);
var Dispatch: IDispatch;
  function GetActiveOleObjectEx(const ClassName: String): Boolean;
  var
    ClassID: TCLSID;
    Unknown: IUnknown;
    AResult: HResult;
  begin
    Result := False;
    ClassID := ProgIDToClassID(ClassName);
    AResult := GetActiveObject(ClassID, nil, Unknown);
    if (Unknown = nil) or not Succeeded(AResult) then Exit;
    AResult := Unknown.QueryInterface(IDispatch, Dispatch);
    if not Succeeded(AResult) then Exit;
    Result := True;
  end;
begin
  fWinPdfCreat := UnAssigned;
  fOk := False;
  if not NewInst then
  begin
    if GetActiveOleObjectEx('PDFCreator.clsPDFCreator') then
    begin
      fWinPdfCreat := Dispatch;
      fOk := True;
      Exit;
    end;
  end;
  try
    fWinPdfCreat := CreateOleObject('PDFCreator.clsPDFCreator');
    fOk := True;
  except
    fWinPdfCreat := UnAssigned;
  end;
  if VarIsEmpty(fWinPdfCreat) then Exit;
  fOk := True;
end;

procedure TpdfCreator.PrepareWorkBook;
begin

end;

end.
