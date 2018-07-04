unit EpurationPieces;

interface

Uses Classes,
     Windows,
     sysutils,
     ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MainEagl,
{$ENDIF}
     HCtrls,
     Hpanel,
     HEnt1,
     HMsgBox,
     Vierge,
     EntGc,
     BTPUtil,
     ExtCtrls,
     SaisUtil,
     UtilPgi,
     uEntCommun,
     UtilTOBPiece,
     UtilEpurePiece,
     UTOB;

type

	TEpurePiece = class (TObject)
  	private
    	fTOBPiece : TOB;
      fresult : integer;
      procedure SupprimePiece;
    public
    	property LaPiece : TOB write fTOBPiece;
      property result : integer read fresult;
    	constructor create;
      destructor destroy; override;
      procedure LanceTraitement;
  end;

function TraitementEpurationPiece (TOBPiece : TOB) : integer;
implementation

function TraitementEpurationPiece (TOBPiece : TOB) : integer;
var EpurePiece : TEpurePiece;
begin
   EpurePiece := TEpurePiece.create;
   EpurePiece.LaPiece := TOBPiece;
   TRY
     EpurePiece.lanceTraitement;
   FINALLY
     result := EpurePiece.Result;
   	 EpurePiece.free;
   END;
end;

{ TEpurePiece }

constructor TEpurePiece.create;
begin

end;

destructor TEpurePiece.destroy;
begin

  inherited;
end;

procedure TEpurePiece.LanceTraitement;
var io   : TIoErr ;
begin
  Io := Transactions (SupprimePiece, 0);
  Case Io of
    oeUnKnown :BEGIN
                fResult := -1;
               END;
  end;
end;

procedure TEpurePiece.SupprimePiece;
begin
	EpureLaPiece (fTOBPiece);
end;

end.
