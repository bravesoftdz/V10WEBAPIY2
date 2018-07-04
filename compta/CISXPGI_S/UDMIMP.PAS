unit UDMIMP;

interface

uses
  Windows, SysUtils, Classes, Controls, Forms,
  HMsgBox, UTob,
  DB, UScript, FileCtrl, stdctrls, HCtrls,
  uDbxDataSet, Variants, ADODB,
  OleServer;

type
  TdmImport = class(TDataModule)
    DataSource1: TDataSource;
    DataSource2: TDataSource;
    DataSource3: TDataSource;
    DataSource4: TDataSource;
    GzImpPar: TADOTable;
    GzImpDelim: TADOTable;
    GzImpReq: TADOTable;
    GzImpCorresp: TADOTable;
    Gzimpdomaine: TADOTable;
    DBGlobal: TADOConnection;
    Cmd: TADOCommand;
    DBGLOBALD: TADOConnection;
    GzImpReqCLE: TWideStringField;
    GzImpReqComment: TWideStringField;
    GzImpReqPARAMETRES: TMemoField;
    GzImpReqTBLCOR: TBlobField;
    GzImpReqClePar: TWideStringField;
    GzImpReqOrigine: TWideStringField;
    GzImpReqDOMAINE: TWideStringField;
    GzImpReqMODIFIABLE: TSmallintField;
    GzImpReqDATEDEMODIF: TDateTimeField;
    GzImpReqTABLE0: TWideStringField;
    GzImpReqTABLE1: TWideStringField;
    GzImpReqTABLE2: TWideStringField;
    GzImpReqTABLE3: TWideStringField;
    GzImpReqTABLE4: TWideStringField;
    GzImpReqTABLE5: TWideStringField;
    GzImpReqTABLE6: TWideStringField;
    GzImpReqTABLE7: TWideStringField;
    GzImpReqTABLE8: TWideStringField;
    GzImpReqTABLE9: TWideStringField;
    ADOTableD: TADOTable;
    CmdDonnee: TADOCommand;
  public
    { Public declarations }
    HdlImpfic : Thandle;
  end;

function InitDB (CurrentPath,CurrentD : String ) : boolean ;
procedure ChargementComboScript (var Comboscript : THVALComboBox; Domaine : string='');
procedure ChargementComboparam (Comboscript : THVALComboBox; Domaine : string='');
function  ExisteScript (Code : string): Boolean;
function  OpenSQLADO (Requete, ConnectString  : string) : TQuery;

Var dmImport : TdmImport ;

implementation

{$R *.DFM}



procedure ActiveDico (CurrentPath, CurrentD : string) ;
var
Rep,FileName       : string;
begin

// domaine d'échange
//   if (VHCX <> nil) and (VHCX^.RepSTD <> '') and ((VHCX^.Domaine = 'E') or (UpperCase(VHCX^.Appelant) = 'ECHANGES PGI'))  then
//      Rep := VHCX^.RepSTD + '\'
//   else
  Rep := ExtractFileDir (Application.ExeName)+'\parametre\';

  FileName := Rep + 'PGZIMPACCESS.mdb';

  // La table des Scripts
  if not FileExists(CurrentPath+'\PGZIMPACCESS.mdb') then
          Copyfile (PChar(FileName), Pchar(CurrentPath+'\PGZIMPACCESS.mdb'),TRUE);
  if not FileExists(CurrentPath+'\Copiefichier.bat') then
          Copyfile (PChar( Rep+'Copiefichier.bat'), Pchar(CurrentPath+'\Copiefichier.bat'),TRUE);
  if not FileExists(CurrentPath+'\RenameFichier.bat') then
          Copyfile (PChar( Rep+'RenameFichier.bat'), Pchar(CurrentPath+'\RenameFichier.bat'),TRUE);

  // La table vide des données
  if not FileExists(CurrentD+'\PGZIMPACCESSD.mdb') then
          Copyfile (PChar(Rep + 'PGZIMPACCESSD.mdb'), Pchar(CurrentD+'\PGZIMPACCESSD.mdb'),TRUE);

end;


function InitDB (CurrentPath,CurrentD : String ) : boolean ;
{$IFNDEF CISXPGI} var Dir, Dir1 : string; {$ENDIF}

procedure Existedirectory (dd :string);
begin
    if not DirectoryExists(dd) then
    begin
{$IFDEF CISXPGI}
              if not DirectoryExists(dd) then CreateDir(dd);
{$ELSE}
       Dir := dd;
       Dir1 := '';
       while  Dir <> '' do
       begin
          if Dir1 = '' then
          Dir1 := ReadTokenPipe (Dir,'\')
          else
          Dir1 := Dir1+'\'+ReadTokenPipe (Dir,'\');
          if not DirectoryExists(Dir1) then
          begin
            if (PGIAsk('Chemin '+ Dir1+ ' inexistant, voulez-vous le créer : ?', 'CISX')= mryes) then
              CreateDir(Dir1)
             else exit;
          end;
       end;
{$ENDIF}
    end;
end;
begin
Result:=FALSE ;
{$IFNDEF CISXPGI}
if CurrentPath ='' then exit ;
if CurrentD ='' then exit;
  try
    Existedirectory (CurrentPath);
    // pour les données
    Existedirectory (CurrentD);
    ChDir(CurrentPath);
  except
    begin
      PGIError('Chemin Invalide','') ;
      exit ;
    end ;
  end ;
{$IFNDEF COMPTA}
ActiveDico  (CurrentPath,CurrentD);
{$ENDIF}

Try
 DMImport.DBGlobal.Connected:=FALSE ;
 DMImport.DBGlobal.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source='+CurrentPath+'\PGZIMPACCESS.mdb'+
 ';Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";'+
 'Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;'+
 'Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";'+
 'Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;'+
 'Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False';
 DMImport.DBGlobal.Connected:=TRUE ;
 except
   On e : exception do
      BEGIN
      PGIError (e.Message,'Erreur de connexion') ;
      exit ;
      END ;
 end;
 try Session.NetFileDir := CurrentPath; except; end;  // pour définir le répertoire PDOXUSRS.NET
{$ELSE}
  if DMImport <> nil then
       exit;
{$ENDIF}
 try
{$IFDEF CISXPGI}
  if DMImport = nil then
     Application.CreateForm(TdmImport, dmImport);
{$ENDIF}

 DMImport.DBGlobalD.Connected:=FALSE ;
 DMImport.DBGlobalD.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source='+CurrentD+'\PGZIMPACCESSD.mdb'+
 ';Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";'+
 'Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;'+
 'Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";'+
 'Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;'+
 'Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False';
 DMImport.DBGlobalD.Connected:=TRUE ;
 except
   On e : exception do
      BEGIN
      PGIError (e.Message,'Erreur de connexion') ;
      exit ;
      END ;
	end;
 DMImport.HdlImpfic := loadlibrary(PChar('IMPFIC7.dll'));

 if DMImport.HdlImpfic = 0 then
     raise Exception.Create('La librairie Impfic7.dll est introuvable.');

CurrentDossier:=CurrentPath ;
CurrentDonnee:=CurrentD;

 // Création de la table s'il n'existe pas
//Futur use   if not Assigned(dmImport) then     dmImport := TdmImport.Create(Application);
Result:=TRUE ;
end;


procedure ChargementComboScript (var Comboscript : THVALComboBox; Domaine : string='');
var
  SelectQL: string;
  Vales   : HTstringList;
  dom     : string;
  QP      : TQuery;
begin
    Vales := HTStringList.Create;
    if not Assigned(dmImport) then
      dmImport := TdmImport.Create(Application);
    if Domaine <> '' then
    SelectQL := ' Where Domaine="'+Domaine+'"'
    else
    SelectQL := '';
    QP := OpenSQLADO ('Select CLE from '+DMImport.GzImpReq.TableName +SelectQL, DMImport.DBGlobal.ConnectionString );
    While  not (QP.Eof) do
    begin
        dom := QP.FindField('CLE').asstring;
        if (Vales.IndexOf(dom) < 0) then
          Vales.add(dom);
        QP.Next;
    end;
    Comboscript.Items := Vales;
    Comboscript.Values := Vales;
    Vales.clear;
    Vales.Free;
     Ferme (QP);
end;

procedure ChargementComboparam (Comboscript : THVALComboBox; Domaine : string='');
var
  Vales   : HTstringList;
  dom     : string;
  QP      : TQuery;
begin
    Vales := HTStringList.Create;
    if not Assigned(dmImport) then
      dmImport := TdmImport.Create(Application);
    QP := OpenSQLADO ('Select * from '+DMImport.GzImpPar.TableName+
    ' Where Domaine="'+Domaine+'"', DMImport.DBGlobal.ConnectionString);
    While  not (QP.Eof) do
    begin
        dom := QP.FindField('TableName').asstring;
        if (Vales.IndexOf(dom) < 0) then
          Vales.add(dom);
        QP.Next;
    end;
    Comboscript.Items := Vales;
    Comboscript.Values := Vales;
    Vales.clear;
    Vales.Free;
    Ferme (QP);
end;

function ExisteScript (Code : string): Boolean;
var
{$IFDEF EAGLCLIENT}
QP    : TOB;
{$ELSE}
QP    : TQuery;
{$ENDIF}
begin
{$IFDEF CISXPGI}
    QP := OpenSQL ('Select CIS_CLE from CPGZIMPREQ Where CIS_CLE="'+Code +'"', TRUE);
{$ELSE}
    QP := OpenSQLADO ('Select CLE from '+DMImport.GzImpReq.TableName +'Where CLE="'+Code +'"', DMImport.DBGlobal.ConnectionString );
{$ENDIF}
    Result := (not QP.EOF);
    Ferme (QP);
end;

function OpenSQLADO (Requete, ConnectString : string) : TQuery;
var
QP : TQuery;
begin
              QP := TQuery.Create(nil);
              QP.ConnectionString := ConnectString;
              QP.SQL.Add (Requete);
              try QP.Open; except; end;
              Result := QP;
end;

end.

