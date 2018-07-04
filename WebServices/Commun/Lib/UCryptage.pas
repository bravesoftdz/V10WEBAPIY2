unit UCryptage;

interface
uses IdHashMessageDigest,IdHash,classes,sysUtils;

function MD5(const Text : string) : string;
function MD5File(const fileName : string) : string;
implementation

function MD5(const Text : string) : string;
var
  md5 : TIdHashMessageDigest5;
begin
  md5 := TIdHashMessageDigest5.Create;
  try
    result := md5.AsHex(md5.HashValue (Text)) ;
  finally
    md5.Free;
  end;
end;

function MD5File(const fileName : string) : string;
 var
   idmd5 : TIdHashMessageDigest5;
   fs : TFileStream;
begin
  idmd5 := TIdHashMessageDigest5.Create;
  fs := TFileStream.Create(fileName, fmOpenRead OR fmShareDenyWrite) ;
  try
    result := idmd5.AsHex(idmd5.HashValue(fs)) ;
  finally
    fs.Free;
    idmd5.Free;
  end;
end;

end.

end.
