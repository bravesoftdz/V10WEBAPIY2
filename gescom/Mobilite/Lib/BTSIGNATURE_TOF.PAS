{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/01/2014
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSIGNATURE ()
Mots clefs ... : TOF;BTSIGNATURE
*****************************************************************}
Unit BTSIGNATURE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
{$ENDIF}
		 AglInit,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Types,
     HTB97,
     UTOF,
     Graphics,
     HRichEdt, HRichOLE,
     JPEG ;

Type
  TOF_BTSIGNATURE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    BREFRESH : TToolbarButton97;
    BVALIDE : TToolbarButton97;
    IMG : THImage;
    fInDrawing : boolean;
    fPrecPoint : TPoint;
    OneTOB : TOB;
    procedure ImGMouseDown(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    procedure ImGMouseMove(Sender: TObject; Shift: TShiftState; X,Y: Integer);
    procedure ImGMouseUp(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    procedure BREFRESHClick (sender : TObject);
    procedure BVALIDECLick (Sender : TObject);
  end ;

   TOF_BTVISUSIGNATURE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
	  CodeAppel : string;
  end ;

function LanceFormeSignature (TOBData : TOB) : boolean;
procedure EnregistreSignature(CodeAppel : string; TOBTT : TOB);
function ExisteSignature(CodeAppel : string) : Boolean;
procedure VisuSignatureClient(CodeAppel : string);

Implementation
uses utilPGI;

procedure VisuSignatureClient(CodeAppel : string);
begin
	AGLLanceFiche('BTP','BTVISUSIGNATURE','','','ACTION=CONSULTATION;CODE='+CodeAppel);
end;

procedure EnregistreSignature(CodeAppel : string; TOBTT : TOB);
begin
  TOBTT.PutValue('LO_TABLEBLOB','APP');
  TOBTT.PutValue('LO_QUALIFIANTBLOB','PHJ');
  TOBTT.PutValue('LO_EMPLOIBLOB','SIG');
  TOBTT.PutValue('LO_IDENTIFIANT',CodeAppel);
  TOBTT.PutValue('LO_RANGBLOB',9);
  TOBTT.SetAllModifie(true);
  TOBTT.InsertDB(nil);
end;

function ExisteSignature(CodeAppel : string) : Boolean;
var QQ: TQuery;
    RR,R1 : boolean;
begin
  // signature
  QQ := OpenSQL('SELECT 1 FROM LIENSOLE WHERE LO_IDENTIFIANT="'+CodeAppel+'" AND '+
  							'LO_TABLEBLOB="APP" AND LO_RANGBLOB=9',True,1,'',true);
  TRY
		RR := not QQ.Eof;
  FINALLY
    ferme (QQ);
  end;
  // commentaire du signataire
  QQ := OpenSQL('SELECT 1 FROM LIENSOLE WHERE LO_IDENTIFIANT="'+CodeAppel+'" AND '+
  							'LO_TABLEBLOB="APP" AND LO_RANGBLOB=10',True,1,'',true);
  TRY
		R1 := not QQ.Eof;
  FINALLY
    ferme (QQ);
  end;
  result := RR OR R1;
end;


function LanceFormeSignature (TOBData : TOB) : boolean;
var OneTOB : TOB;
begin
	OneTOB := TOB.Create ('UNETOB',nil,-1);
  TRY
    OneTOB.AddChampSupValeur('RESULT','-');
    OneTOB.AddChampSup ('SIGNATURE',false);
    OneTOB.AddChampSupValeur ('NOM','');
    TheTOB := OneTOB;
    AGLLanceFiche('BTP','BTSIGNATURE','','','');
    TheTOB := nil;
    Result := (OneTOB.GetValue('RESULT')='X');
    if result then
    begin
      TOBData.Putvalue('LO_OBJET',OneTOB.GetValue('SIGNATURE'));
  		TOBData.PutValue('LO_LIBELLE',OneTOB.GetValue('NOM'));
    end;
  FINALLY
  	OneTOB.free;
  end;
end;


procedure TOF_BTSIGNATURE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSIGNATURE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSIGNATURE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTSIGNATURE.OnLoad ;
begin
  Inherited ;
	IMG.Canvas.FillRect(IMG.ClientRect);
end ;

procedure TOF_BTSIGNATURE.OnArgument (S : String ) ;
begin
  Inherited ;
  OneTOB := LaTOB;
  BREFRESH := TToolbarButton97 (GetControl('BREFRESH'));
  BREFRESH.OnClick := BREFRESHClick;
  BVALIDE := TToolbarButton97 (GetControl('BVALIDER_'));
  BVALIDE.OnClick := BVALIDECLick;
	IMG := THImage(GetControl('IMG'));
  IMG.OnMouseDown := ImGMouseDown;
  IMG.OnMouseUp := ImGMouseUp;
  IMG.OnMouseMove := ImGMouseMove;
  IMG.Canvas.Pen.Width := 6;
  fInDrawing := false;
end ;

procedure TOF_BTSIGNATURE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSIGNATURE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSIGNATURE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSIGNATURE.ImGMouseDown(Sender: TObject;Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbLeft then
  begin
    fPrecPoint := Point(X,Y);
    fInDrawing := True;
    ImG.Canvas.MoveTo(X,Y);
  end;
end;

procedure TOF_BTSIGNATURE.ImGMouseMove(Sender: TObject; Shift: TShiftState;X, Y: Integer);
begin
	if fInDrawing then
  begin
    Img.Canvas.LineTo(X,Y);
    ImG.Invalidate;
  end;
end;

procedure TOF_BTSIGNATURE.ImGMouseUp(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbLeft then
  begin
  	fInDrawing := false;
  end;
end;


procedure TOF_BTSIGNATURE.BREFRESHClick(sender: TObject);
begin
	IMG.Canvas.FillRect(IMG.ClientRect);
end;

procedure TOF_BTSIGNATURE.BVALIDECLick(Sender: TObject);
var TTBMP: TMemoryStream;
		StImage : string;
    XX : TJPEGImage;
    YY : TBitMap;
begin
  TTBMP := TMemoryStream.Create;
  XX := TJPEGImage.Create;
  TRY
    YY := IMG.Picture.Bitmap;
    BitmapToJPeg (YY,XX);
    XX.SaveToStream (TTBMP);
    SetLength(StImage,TTBMP.Size);
    TTBMP.Seek(0,0);
    TTBMP.Read(pchar(StImage)^,TTBMP.Size);
    OneTOB.PutValue('SIGNATURE',StImage);
    OneTOB.PutValue('NOM',GetControlText('NOM'));
    OneTOB.PutValue('RESULT','X');
  FINALLY
  	TTBMP.Free;
    XX.Free;
  end;
end;

{ TOF_BTVISUSIGNATURE }

procedure TOF_BTVISUSIGNATURE.OnArgument(S: String);
var critere,St : string;
		x : Integer;
    Champmul,ValMul : string;
    QQ : TQuery;
begin
  inherited;
  St := S;
  repeat
    Critere := Trim(ReadTokenSt(st));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'CODE' then CodeAppel := ValMul;
      end;
    end;
  until Critere = '';
	if CodeAppel <> '' then
  begin
		QQ := OpenSQL('SELECT LO_LIBELLE,LO_OBJET FROM LIENSOLE WHERE '+
    						  'LO_IDENTIFIANT="'+CodeAppel+'" AND '+
                  'LO_TABLEBLOB="APP" AND LO_RANGBLOB=9',True,1,'',True);
    if not QQ.eof then
    begin
			THEdit(GetControl('NOM')).text := QQ.FindField('LO_LIBELLE').AsString;
    	LoadBitMapFromChamp(QQ, 'LO_OBJET', THImage(Getcontrol('IMAGE')), true);
    end;
    ferme (QQ);
    // 
		QQ := OpenSQL('SELECT LO_OBJET FROM LIENSOLE WHERE '+
    						  'LO_IDENTIFIANT="'+CodeAppel+'" AND '+
                  'LO_TABLEBLOB="APP" AND LO_RANGBLOB=10',True,1,'',True);
    if not QQ.eof then
    begin
      St := QQ.fields[0].AsString;
      StringToRich(THRichEditOLE (GetControl('COMMENT')), st);
    end;
    ferme (QQ);
  end;
end;

procedure TOF_BTVISUSIGNATURE.OnCancel;
begin
  inherited;

end;

procedure TOF_BTVISUSIGNATURE.OnClose;
begin
  inherited;

end;

procedure TOF_BTVISUSIGNATURE.OnDelete;
begin
  inherited;

end;

procedure TOF_BTVISUSIGNATURE.OnDisplay;
begin
  inherited;

end;

procedure TOF_BTVISUSIGNATURE.OnLoad;
begin
  inherited;

end;

procedure TOF_BTVISUSIGNATURE.OnNew;
begin
  inherited;

end;

procedure TOF_BTVISUSIGNATURE.OnUpdate;
begin
  inherited;

end;

Initialization
  registerclasses ( [ TOF_BTSIGNATURE ] ) ;
  registerclasses ( [ TOF_BTVISUSIGNATURE ] ) ;
end.
