unit UPrintScreen;

interface
uses StdCtrls,
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
     Windows,
     Graphics,
     JPEG,
     EdtREtat ;

procedure PrintThisScreen ( OneForm : TForm ; NomFiche : String='');

implementation

procedure PrintThisScreen ( OneForm : TForm; NomFiche : String='');
var
		FDC : HDC;
    XX : TBitmap;
    JPGIMG : TJPEGImage;
    OneTOB,OneDetail : TOB;
    TheImageSt,StName : string;
    MEMSTREAM : TMemoryStream;
    StSql,Identifiant : string;
begin
  OneTOB := TOB.Create ('UNE TOB',nil,-1);
  OneTOB.AddChampSupvaleur ('UN CHAMP','');
  OneDetail := TOB.Create ('LIENSOLE',OneTOB,-1);
  Identifiant := AglGetGuid;
  StName := NomFiche;
  if StName = '' then StName := OneForm.Caption;
  OneDetail.PutValue('LO_IDENTIFIANT',Identifiant);
  OneDetail.PutValue('LO_RANGBLOB',0);
  OneDetail.PutValue('LO_CREATEUR',V_PGI.User);
    OneDetail.PutValue('LO_LIBELLE',stName);
  JPGIMG := TJPEGImage.Create;
  XX := TBitmap.create;
  MEMSTREAM := TMemoryStream.Create;
  TRY
    FDC := GetDC(OneForm.Handle);
    XX.Width := OneForm.ClientWidth;
    XX.Height := OneForm.ClientHeight;
    BitBlt (XX.Canvas.Handle,0,0,XX.Width,XX.Height,FDC,0,0,SRCCOPY);

    BitmapToJPeg (XX,JPGIMG); // passage du bitmap en JPEG
    JPGIMG.SaveToStream (MEMSTREAM); // Passage du JPEG EN MEMORY STREAL
    // PASSAGE DU MEM STREAM EN CHAINE DE CARACTERES
    SetLength(TheImageSt,MEMSTREAM.Size);
    MEMSTREAM.Seek(0,0);
    MEMSTREAM.Read(pchar(TheImageSt)^,MEMSTREAM.Size);
    // -----
    OneDetail.PutValue('LO_OBJET',TheImageSt);
    OneDetail.InsertOrUpdateDB(false);
    StSql := 'SELECT * FROM LIENSOLE WHERE LO_IDENTIFIANT="'+Identifiant+'" AND Lo_RANGBLOB=0';
    LanceEtat  ('E','YYY','BPS', true,false,False,nil,StSql,StName,false);
    ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_IDENTIFIANT="'+identifiant+'"');
  finally
    OneTOB.Free;
    XX.Free;
    JPGIMG.free;
    MEMSTREAM.Free;
  end;
end;

end.
