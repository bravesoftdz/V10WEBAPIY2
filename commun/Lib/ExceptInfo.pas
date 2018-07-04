unit ExceptInfo;

interface

uses
  Forms,
  SysUtils,
  Windows,
  hmsgbox,
  Clipbrd,
  Hent1,
  HDebug,
  Controls ;

type
  TCacheMisere = class
  protected
    Enabled : boolean ;
    CacheViolation : Boolean ;
    InException : Boolean ;
    ViolationCount : Integer ;
    procedure ApplicationException(Sender: TObject; E: Exception);
  end;

// Sert a pouvoir modifier la valeur de CacheViolation.
var CacheMisere : TCacheMisere;


implementation

procedure TCacheMisere.ApplicationException(Sender: TObject; E: Exception);
const None : String = 'Aucun' ;
var
 ErrorText, ErrorAdress : String;
 Inviolation : Boolean ;
 function GetActiveFormName : String ;
 begin
    if Assigned(Screen.ActiveForm) then
       Result:= Screen.ActiveForm.Name
    else
       Result := None ;
 end;
 function GetActiveControlName : String ;
 begin
    if Assigned(Screen.ActiveControl) then
       Result:= Screen.ActiveControl.Name
    else
       Result := None ;
 end;
 function GetErrorAdress : String ;
 begin
   try
      Result := IntToHex(Integer(ExceptAddr), 8) ;
      Clipboard.AsText := Result ;
   except
      Result := None ;
   end;
 end;
begin
 Debug('-- APPLICATION.ONEXCEPTION');
 Debug('-- ERREUR GRAVE : '+e.Message );
 Try
  ErrorAdress := GetErrorAdress ;
  ErrorText:= 'Informations sur l''erreur : ' + #13#10
   + ExtractFileName(Application.ExeName) + ' ' + V_PGI.NumVersion+'.'+V_PGI.NumBuild + ' ('+AGLNumVersion+')'+ #13#10
   + 'Heure = ' + DateTimeToStr(Now) + #13#10
   + 'Texte = ' + E.Message + #13#10
   + 'Classe d''exception = ' + E.ClassName + #13#10
   + 'Adresse = ' + ErrorAdress + #13#10
   + 'Fenêtre Active = ' + GetActiveFormName + #13#10
   + 'Control Actif = ' + GetActiveControlName ;
  Debug('-- ERREUR INFO : '+ErrorText );
  Inviolation := ( E.ClassName = 'EAccessViolation' ) ;
  if not Inviolation then ViolationCount := 0 ;
  // Si violation d'acces ... on masque si CacheViolation = TRUE et non V_PGI.SAV
  if not V_PGI.SAV
  AND CacheViolation
  and Inviolation then
  begin
     inc(ViolationCount) ;
     if ViolationCount < 3 then
        exit ;
  end;
  // sinon violation d'acces
  if InException then exit ;
  InException := True ;
//   ( pos ('Violation',E.Message) > 0) then exit ;
  if Inviolation then ErrorText := ErrorText+#13#10+#13#10+'L''application va se terminer ! '+#13#10 ;
  PGIBox(ErrorText) ;
  if Inviolation then Application.Terminate ;
 Except
   on e2: Exception do
   begin
     ErrorText := 'Error in default exception handler.'+#13#10+'Erreur Principale : '+E.Message+'('+ErrorAdress+')'+#13#10+'Erreur Secondaire : '+E2.Message ;
     Debug('-- ERREUR INFO : '+ErrorText );
     PGIError(ErrorText);
     Application.Terminate ;
   end;
 End;
 InException := False ;
end;

Function FileSize(FileName : String) : Int64;
var
  SearchRec : TSearchRec;
begin
  if SysUtils.FindFirst(FileName, faAnyFile, SearchRec ) = 0 then
  begin
     // if found
     Result := Int64(SearchRec.FindData.nFileSizeHigh) shl Int64(32) +    // calculate the size
               Int64(SearchREc.FindData.nFileSizeLow);
     SysUtils.FindClose(SearchRec);                                       // close the find
  end
  else
     Result := 0;
end;


initialization
  CacheMisere := TCacheMisere.Create;
  CacheMisere.InException := False ;
  CacheMisere.CacheViolation := False ;
  CacheMisere.ViolationCount := 0 ; 
  CacheMisere.Enabled := FileExists(IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName))+'pgiacces.dat') ;
  if CacheMisere.Enabled then
     CacheMisere.CacheViolation := FileSize(IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName))+'pgiacces.dat') > 0 ;
  If CacheMisere.Enabled then
     Application.OnException := CacheMisere.ApplicationException;
finalization
//  Application.OnException := nil ;
  CacheMisere.Free;
end.


