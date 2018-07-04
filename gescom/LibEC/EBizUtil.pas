unit EBizUtil;

interface
uses Classes, UTOB{, wtobdebug_tof};
{$IFNDEF EAGLCLIENT}
type TOB_2 = class(TOB)
     private
       function YouzTob(T : TOB) : integer;

     public
       procedure LoadFromXMLStream(FStream : TStream);
       procedure LoadFromBinStream(FStream : TStream);
       procedure LoadFromXMLFile(FName : String);

     end;

function TOBtoString(Entree : TOB; Zipped : boolean = false) : String;
function TOBtoBINString(Entree : TOB; Zipped : boolean = false) : String;
function TOBtoXMLString(Entree : TOB; Zipped : boolean = false) : String;
function TOBtoXMLStringSupp(Entree : TOB; Zipped : boolean = false; Entete : boolean = True; Supp : boolean = False) : String;
procedure BINStringToTOB(const Entree : String; Sortie : TOB; Zipped : boolean = false);
procedure XMLStringToTOB(const Entree : String; Sortie : TOB; Zipped : boolean = false);
procedure XXXFileToTOB(FichierEntree : String; Sortie : TOB; BinMode : boolean; zipped : boolean = false);
procedure TOBToXXXFile(Entree : TOB; FichierSortie : String; BinMode : boolean; zipped : boolean = false);
{$ENDIF}
function ChampType(FieldS : String) : String;
//procedure ScanRates(T : TOB);
{$IFNDEF EAGLCLIENT}
procedure TAConnecte(FConnectedeuxfois,FichierINI, NomSoc : OleVariant; ChargeProut : boolean = true);
procedure TADeConnecte;
procedure TAIntegreDataFichier(NomFichier: OleVariant; BinMode : Boolean; zipped : boolean = false);
{$ENDIF}
implementation
uses {$IFNDEF EAGLCLIENT}controls, HZLib, MajTable, Math, Dialogs, sysutils, UxmlUtils, EasyMSMQ,{$ENDIF}
     HEnt1, HCtrls, Ent1, EntGC
    ,CbpMCD
    ,CbpEnumerator
     ;

{$IFNDEF EAGLCLIENT}
function TOB_2.YouzTob(T : TOB) : integer;
begin
  T.ChangeParent(Self, -1);
  T.SetAllModifie(true);
  result := 0;
end;

procedure TOB_2.LoadFromXMLStream(FStream : TStream);
begin
  TOBLoadFromXMLStream(FStream, YouzTob);
end;

procedure TOB_2.LoadFromBinStream(FStream : TStream);
begin
  TOBLoadFromBinStream(FStream, YouzTob, nil);
end;

procedure TOB_2.LoadFromXMLFile(FName : String);
begin
  TOBLoadFromXMLFile(FName, YouzTob);
end;

procedure ZipStream(var S : TStringStream);
var PBuf : Pointer;
    PSize : Integer;
begin
    ZipBuffer(Addr(S.DataString[1]), S.Size, PBuf, PSize);
    S.Free;
    S := TStringStream.Create('');
    S.Write(PBuf^, PSize);
    FreeMem(PBuf);
end;

procedure UnZipStream(var S : TStringStream);
var PBuf : Pointer;
    PSize : Integer;
begin
    UnZipBuffer(Addr(S.DataString[1]), S.Size, 0, PBuf, PSize);
    S.Free;
    S := TStringStream.Create('');
    S.Write(PBuf^, PSize);
    FreeMem(PBuf);
end;

//procedure EcritDansFichier(mess, fichier: string);
//var
//  fic: TextFile;
//begin
//  assignFile(fic, fichier);
//  append(fic);
//  writeln(fic, mess);
//  close(fic);
//end;


function TOBtoString(Entree : TOB; Zipped : boolean = false) : String;
var Sortie : TStringStream;
begin
  Sortie := TStringStream.Create(Entree.SaveToBuffer(false, false, ''));

  if Zipped then ZipStream(Sortie);

  result := Sortie.DataString;
  Sortie.Free;
end;

function TOBtoBINString(Entree : TOB; Zipped : boolean = false) : String;
var Sortie : TStringStream;
begin
  Sortie := TStringStream.Create('');
  Entree.SaveToBinStream(Sortie, false, true, false, false);

  if Zipped then ZipStream(Sortie);

  result := Sortie.DataString;
  Sortie.Free;
end;

function TOBtoXMLString(Entree : TOB; Zipped : boolean = false) : String;
var Sortie : TStringStream;
begin
  Sortie := TStringStream.Create('');
  Entree.SaveToXmlStream(Sortie, true, false);

  if Zipped then ZipStream(Sortie);

  result := Sortie.DataString;
  Sortie.Free;
end;

function TOBtoXMLStringSupp(Entree : TOB; Zipped : boolean = false; Entete : boolean = True; Supp : boolean = False) : String;
var Sortie : TStringStream;
begin
  Sortie := TStringStream.Create('');
  Entree.SaveToXmlStream(Sortie, Entete, Supp);

  if Zipped then ZipStream(Sortie);

  result := Sortie.DataString;
  Sortie.Free;
end;

// La tob de sortie est déjà créée (et vide)
procedure BINStringToTOB(const Entree : String; Sortie : TOB; Zipped : boolean = false);
var StrStr : TStringStream;
begin
  if Entree = '' then exit;
  StrStr := TStringStream.Create(Entree);

  if Zipped then UnZipStream(StrStr);

  StrStr.Seek(0, soFromBeginning);
  TOB_2(Sortie).LoadFromBinStream(StrStr);
  StrStr.Free;
end;

// La tob de sortie est déjà créée (et vide)
procedure XMLStringToTOB(const Entree : String; Sortie : TOB; Zipped : boolean = false);
var StrStr : TStringStream;
begin
  if Entree = '' then exit;
  StrStr := TStringStream.Create(Entree);

  if Zipped then UnZipStream(StrStr);

  StrStr.Seek(0, soFromBeginning);
  TOB_2(Sortie).LoadFromXMLStream(StrStr);
  StrStr.Free;
end;

{procedure XMLFileToTOB(FichierEntree : String; Sortie : TOB);
begin
TOB_2(Sortie).LoadFromXMLFile(FichierEntree);
end;}

procedure XXXFileToTOB(FichierEntree : String; Sortie : TOB; BinMode : boolean; zipped : boolean = false);
var XeuMeuLeu : String;
    FIn : TFileStream;
begin
  FIn := TFileStream.Create(FichierEntree, fmOpenRead);
  SetLength(XeuMeuLeu, FIn.Size);
  FIn.Read(XeuMeuLeu[1], FIn.Size);
  FIn.Free;

  if BinMode then BINStringToTOB(XeuMeuLeu, Sortie, zipped)
             else XMLStringToTOB(XeuMeuLeu, Sortie, zipped);
end;

procedure TOBToXXXFile(Entree : TOB; FichierSortie : String; BinMode : boolean; zipped : boolean = false);
var XeuMeuLeu : String;
    FOut : TFileStream;
begin
  if BinMode then XeuMeuLeu := TOBToBINString(Entree, zipped)
             else XeuMeuLeu := TOBToXMLString(Entree, zipped);

  FOut := TFileStream.Create(FichierSortie, fmCreate);
  FOut.Write(XeuMeuLeu[1], SizeOf(XeuMeuLeu[1])*Length(XeuMeuLeu));
  FOut.Free;
end;
{$ENDIF}

function ChampType(FieldS : String) : String;
var iTable, iChamp : integer;
    PrefX : String;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
	NomChamps : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  result := '';
  PrefX := Copy(FieldS, 1, Pos('_', FieldS)-1);
  Table := Mcd.GetTable(Mcd.PrefixeToTable(PrefX));
  FieldList := Table.Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
  	NomChamps := (FieldList.Current as IFieldCOM).name;
  	if NomChamps = FieldS then
    begin
      result := (FieldList.Current as IFieldCOM).Tipe;
      break;
    end;
  end;
end;

(*procedure ScanRates(T : TOB);
var iLig, iChamp : integer;
const NbDecim = 15;
begin
  for iLig := 0 to T.Detail.Count-1 do
  begin
    for iChamp := 1 to T.Detail[iLig].NbChamps do
      if (ChampType(T.Detail[iLig].GetNomChamp(iChamp)) = 'RATE')
      or (ChampType(T.Detail[iLig].GetNomChamp(iChamp)) = 'DOUBLE') then
        T.Detail[iLig].Valeurs[iChamp] := Int(T.Detail[iLig].Valeurs[iChamp] * Power(10, NbDecim)) / Power(10, NbDecim);

    if T.Detail[iLig].Detail.Count > 0 then ScanRates(T.Detail[iLig]);
  end;
end;*)
{$IFNDEF EAGLCLIENT}
procedure TAConnecte(FConnectedeuxfois,FichierINI, NomSoc : OleVariant; ChargeProut : boolean = true);
begin
  //InitLaVariableHalley; fait dans la phase initialization
  //EcritDansFichier('Après InitLaVariableHalley','logconnecthalley.log');
  //InitLaVariableGC; fait dans la phase initialization
  //EcritDansFichier('Après InitLaVariableGC','logconnecthalley.log');
  HalSocIni := FichierIni;
  V_PGI.User:='000' ; V_PGI.UserName:='00000000' ; V_PGI.PassWord:='000' ;

// Modif C TOUDIC Pages Jaunes
  V_PGI.NoLock := TRUE;
  {PG}
  //V_PGI.OracleAutoCommit := TRUE;

//  V_PGI.User:='IFP' ; V_PGI.UserName:='IFP' ; V_PGI.PassWord:='305197FF' ;
  if ConnecteHalley(NomSoc,False,{Nil}@ChargeMagHalley,Nil,Nil,nil) then // ChargeMagHalley rajout 12/11/02 pour ImportCCAuto
   begin
     //EcritDansFichier('Après ConnecteHalley','logconnecthalley.log');
     V_PGI.OKOuvert:=TRUE ; V_PGI.CurrentAlias:=NomSoc ;
     //if ChargeProut then
     //begin
     //    ChargeTVATPF ;
     //    ChargeSocieteHalley;
     //    ChargeDescriGC;
     //end;
             //if FConnectedeuxfois=true then
             //begin
             {     if VH_GC.TOBMEA<>Nil then VH_GC.TOBMEA.Free ;
                  if VH_GC.TOBParPiece<>Nil then VH_GC.TOBParPiece.Free ;
                  if VH_GC.TOBParPieceComp<>Nil then VH_GC.TOBParPieceComp.Free ;
                  if VH_GC.TOBParPieceDomaine<>Nil then VH_GC.TOBParPieceDomaine.Free ;
                  VH_GC.TOBMEA:=TOB.Create('',Nil,-1) ; VH_GC.TOBMEA.LoadDetailDB('MEA','','',Nil,True) ;
                  VH_GC.TOBParPieceDomaine:=TOB.Create('',Nil,-1) ; VH_GC.TOBParPieceDomaine.LoadDetailDB('DOMAINEPIECE','','',Nil,True) ;
                  VH_GC.TOBParPieceComp:=TOB.Create('',Nil,-1) ; VH_GC.TOBParPieceComp.LoadDetailDB('PARPIECECOMPL','','',Nil,True) ;
                  VH_GC.TOBParPiece:=TOB.Create('',Nil,-1) ; VH_GC.TOBParPiece.LoadDetailDB('PARPIECE','','',Nil,True) ;
                  if VH_GC.TOBParPiece.Detail.Count<=0 then Exit ;   }
             //end;

   end;
end;

procedure TADeConnecte;
begin
  DeconnecteHalley ;
  V_PGI.OKOuvert:=FALSE ;
end;

procedure TAIntegreDataFichier(NomFichier: OleVariant; BinMode : Boolean; zipped : boolean = false);
Var TOBG : TOB ;
begin
  TOBG:=TOB.Create('',Nil,-1) ;
  XXXFileToTOB(NomFichier,TOBG, BinMode, zipped);
  TOBG.InsertOrUpdateDB ;
  TOBG.Free ;
end;
{$ENDIF}

end.
