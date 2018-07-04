unit GCSoucheUtil;

// Utilisation des souches pour la GC (Type = GES)

interface

function GCSoucheGetNext (const CodeSouche : String; Combien : integer = 1) : Integer;
function GCSoucheReadCurrent (const CodeSouche : string) : integer;
procedure GcSoucheSetValue (const CodeSouche : string; Compteur : integer);
function GCInitSequenceSouche (bInit : boolean = false) : integer;
// CRM_20090924_MNG_FQ;500;16769
procedure GcSoucheInitValue (const CodeSouche : string; Compteur : integer);
procedure GCSoucheDelete (const CodeSouche : string);

implementation

uses
  uTob,
  {$IFNDEF EAGLCLIENT}
    uDbxDataSet,
  {$ENDIF}
  HCtrls,
  cbpDBSequences,
  Hent1,
  ySouche_abs,
  yServiceManager,
  SysUtils, DB;

Const
  TypeSouche = 'GES';

function GCSoucheGetNext (const CodeSouche : String; Combien : integer = 1): Integer;
var
  MaSouche : IySouche;
begin
  MaSouche := TYServiceManager.SoucheManager;
  Result := MaSouche.GetNext (TypeSouche, CodeSouche, Combien)
end;

function GCSoucheReadCurrent (const CodeSouche : String): Integer;
var
  MaSouche : IySouche;
begin
  MaSouche := TYServiceManager.SoucheManager;
  Result := MaSouche.ReadCurrent (TypeSouche, CodeSouche)
end;

procedure GCSoucheSetValue (const CodeSouche : string; Compteur : integer);
var
  masouche : IySouche;
begin
  { CRM_20090721_MNG_FQ;010;17361_DEB }
  try
    MaSouche := TYServiceManager.SoucheManager;
    maSouche.SetValue (TypeSouche, CodeSouche, Compteur);
  except
    on e : Exception do
      raise(exception.create(e.Message));
  end;
 { CRM_20090721_MNG_FQ;010;17361_FIN }
end;

function GCInitSequenceSouche (bInit : boolean = false) : integer;
var
  TobSouche : TOB;
  iSouche : integer;
  CodeSouche : string;
begin
  // CRM_20090721_MNG_FQ;010;17361
  result:=0;
  TobSouche := Tob.Create ('', nil, -1);
  try
    TobSouche.LoadDetailFromSQL('SELECT SH_SOUCHE FROM SOUCHE WHERE SH_TYPE="GES"');
    for iSouche := 0 to TobSouche.Detail.Count - 1 do
    begin
      CodeSouche := TobSouche.Detail[iSouche].GetString ('SH_SOUCHE');
      if bInit then
        begin
        { CRM_20090721_MNG_FQ;010;17361_DEB }
          try
            // CRM_20090924_MNG_FQ;500;16769
            //GCSoucheSetValue (CodeSouche, 0);
            GCSoucheInitValue (CodeSouche, 0) ;
          except
          on e : Exception do
            raise(exception.create(e.Message));
          end;
        { CRM_20090721_MNG_FQ;010;17361_FIN }
        end
      else
        GCSoucheReadCurrent (CodeSouche);
    end;
    Result := TobSouche.Detail.Count;
  finally
    TobSouche.Free;
  end;
end;

{ CRM_20090924_MNG_FQ;500;16769_DEB }
procedure GCSoucheInitValue (const CodeSouche : string; Compteur : integer);
var
  masouche : IySouche;
begin
  MaSouche := TYServiceManager.SoucheManager;
  try
    if maSouche.Exists (TypeSouche, CodeSouche) then
      maSouche.Remove (TypeSouche, CodeSouche);
    maSouche.CreateFirst (TypeSouche, CodeSouche, Compteur);
  except
    on e : Exception do
      raise(exception.create(e.Message));
  end;
end;
{ CRM_20090924_MNG_FQ;500;16769_FIN }

procedure GCSoucheDelete (const CodeSouche : string);
var
  masouche : IySouche;
begin
  MaSouche := TYServiceManager.SoucheManager;
  try
    maSouche.Remove (TypeSouche, CodeSouche);
  except
    on e : Exception do
      raise(exception.create(e.Message));
  end;
end;

end.
