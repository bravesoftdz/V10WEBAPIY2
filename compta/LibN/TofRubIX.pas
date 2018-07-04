unit TofRubIX;

interface

uses Classes, StdCtrls, SysUtils, dbtables, Dialogs,
     UTof, HCtrls, HEnt1, Mul, HTB97, hmsgbox;

type
  TOF_RubImpExp = class(TOF)
  private
    btnExport: TToolbarButton97;
    procedure btnExportClick(Sender: TObject) ;
    procedure RubExport(Strs: TStrings; FileName: string);
    procedure ExportAscii(sql: String; FileName: string);
  public
    procedure OnArgument(Arguments : string) ; override ;
    procedure OnUpdate ; override ;
    procedure OnLoad ; override ;
  end ;

implementation


{ TOF_RubImpExp }

procedure TOF_RubImpExp.btnExportClick(Sender: TObject);
var sd: TSaveDialog;
    OldDefaultExt, OldFilter: string;
begin
  sd:=TFMul(Ecran).SD;
  try
    OldDefaultExt:=sd.DefaultExt;
    OldFilter:=sd.Filter;
    sd.DefaultExt:='TXT';
    sd.Filter:='Fichier Texte (*.txt)|*.txt';
    if sd.Execute then RubExport(TFMul(Ecran).Z_SQL.Lines, sd.FileName);
  finally
    sd.DefaultExt:=OldDefaultExt;
    sd.Filter:=OldFilter;
  end;
end;

procedure TOF_RubImpExp.RubExport(Strs: TStrings; FileName: string);
var i, j: integer;
    s: string;
BEGIN
s:='';
for i:=0 to Strs.Count-1 do
  s:=s+Strs.Strings[i]+' ';
s:=uppercase(s);
i:=pos('SELECT',s)+6 ;
j:=pos('FROM',s)-i;
System.Delete(s,i,j);
System.Insert(' * ', s, i);
ExportAscii(s,FileName);
END;

procedure TOF_RubImpExp.ExportAscii(sql: String; FileName: string);
var f : TextFile ;
    Qr: TQuery;
    ch, ln: string;
    i: integer;
BEGIN
{$I-}AssignFile(f,FileName) ; rewrite(f) ;{$I+}
if IoResult <> 0 then begin
  HShowMessage('0;Erreur;Impossible de créer le fichier;W;O;O;O','','') ;
  Exit;
end;

ch:='|';
try
  SourisSablier;
  Qr:=OpenSql(sql, true);
  while not Qr.Eof do begin
    ln:='';
    for i:=0 to Qr.FieldCount-1 do begin
      ln:=ln+Qr.Fields[i].DisplayText;
      if i<Qr.FieldCount then ln:=ln+ch;
    end;
    WriteLn(f,ln) ;
    Qr.Next;
  end;
finally
  Ferme(Qr);
  CloseFile(f) ;
  SourisNormale;
end;
END;

procedure TOF_RubImpExp.OnArgument(Arguments: string);
begin
  btnExport:=TToolbarButton97(GetControl('BOuvrir'));
  if btnExport<>nil then begin
    btnExport.Caption:='Exportation';
    btnExport.Hint:='Exportation des enregistements sélectionnés';
    btnExport.OnClick:=btnExportClick;
  end;
end;

procedure TOF_RubImpExp.OnLoad;
begin

end;

procedure TOF_RubImpExp.OnUpdate;
begin

end;


initialization
RegisterClasses([TOF_RubImpExp]);

end.
