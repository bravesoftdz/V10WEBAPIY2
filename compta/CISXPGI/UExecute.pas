unit UExecute;

interface

uses
  Windows, SysUtils, Classes, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, UScript, HTB97, HMsgBox, HQRY,
  UTOB,
  ed_tools, Hent1, Math,
  HCtrls, DB,
{$IFDEF CISXPGI}
   UControlParam,
{$ENDIF}
{$IFDEF EAGLCLIENT}
   UScriptTob,
{$ELSE}
   Fe_main,
{$ENDIF}
  uDbxDataSet, Variants, ADODB,
  HStatus;


type
  TCallBackDlg = class(TForm)
    Bevel2: TBevel;
    Bevel6: TBevel;
    Bevel7: TBevel;
    Bevel5: TBevel;
    Bevel4: TBevel;
    Label1: TLabel;
    Label2: TLabel;
    lblLu: TLabel;
    lblSel: TLabel;
    Label5: TLabel;
    lblAnomalies: TLabel;
    lblPerf: TLabel;
    Label3: TLabel;
    PaintBox1: TPaintBox;
    Bevel1: TBevel;
    Label4: TLabel;
    lblInfoSoc: TLabel;
    Label6: TLabel;
    Timer1: TTimer;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    bDefaire: TToolbarButton97;
    Binsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BImprimer: TToolbarButton97;
    procedure FormHide(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BFermeClick(Sender: TObject);
  private
    { Déclarations privées }
    Count                         : Longint;
    prevEtape                     : Integer;
    Tick                          : integer;
    FAnomalies, FPrevAnomalies    : Integer;
    FEnrLu, FPrevEnrLu            : Integer;
    FEnrSel, FPrevEnrSel          : Integer;
    InitProgress                  : integer;
    OkControle                    : Boolean;
    procedure SetEnrLu(AValue:integer);
	procedure SetEnrSel(AValue:integer);
	procedure SetAnomalies(AValue:integer);


  public
    { Déclarations publiques }
    bAbort           : Boolean;
    CountLu          : Integer;
    sDzMC1Par        : String;
    bClosing         : Boolean;
    function LaCallBack(Sender:TObject; AEtape:Integer;
                   FBuffer:PChar; bSelected:Boolean;var OutputList:TOutputList)
						: integer; stdcall;
    property EnrLu : Integer write SetEnrLu;
    property EnrSel : Integer write SetEnrSel;
    property Anomalies : Integer write SetAnomalies;
    function InsertRecord(const FScript:TScript; FBuffer:PChar) : Boolean;
    procedure Free;
    function CallbackScrut(ANumEnr: Longint; AEtape: Integer): integer; stdcall;
    function ExisteEnreg(var Code : string; table,field,ParName : string) : boolean;
    function ControleCorresp (S : string; ii : integer; Scr : TScript): string;
  end;

var
  CallBackDlg: TCallBackDlg;

function EcritureFichierExport(var F : TextFile; Sortie, Requete,Domaine  : string; var DatedebutExo, DatefinExo : TDateTime; Lien : string=''; Where : string=''; ScriptName : string=''; Ordre : string=''; Lepays : string='FRA'; PlusLien : TOB=nil ; Tbp : TOB=nil): string;
{$IFDEF CISXPGI}
function EcritureDanslaBase(Sortie, Requete,Domaine  : string; var TEcr, TReleve : TOB; var CpteGeneral, RefPointage : string; var NumReleve : integer; var TPControle : TSVControle; Where : string=''; Ordre : string=''; IntegratioReleve :Boolean=TRUE; IntegrationEtebac:Boolean=FALSE): string;
{$ENDIF}
procedure ExportFichierText(var F : TextFile; FichS : string);
Function RenseigneDomaine  (Cle : string) :string;

implementation

uses
  {$IFDEF CISXPGI}
   UtilPGI,
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ELSE}
   ent1,
  {$ENDIF MODENT1}
  {$ENDIF}

  UDMIMP;


{$R *.DFM}
Function Format_Date(Dat : String) : TDateTime ;
var An    : String ;
    Y,M,J : Word ;
    DD,D1 : string;
BEGIN
Result := iDate1900 ;
if Trim(Dat)='' then Exit ;
DD := ''; D1 := Dat;
While D1 <> ''  do
     DD := DD + ReadTokenPipe(D1,'/');
if DD = '' then DD := Dat;
An:=Copy(DD,5,4) ;
Y:=Round(Valeur(An)) ;
M:=Round(Valeur(Copy(DD,3,2))) ;
J:=Round(Valeur(Copy(DD,1,2))) ;
If (J=0) Or (M=0) Or (Y=0) Then Exit ;
If j in [1..31]=FALSE Then Exit ;
If m in [1..12]=FALSE Then Exit ;
Result:=EncodeDate(Y,M,J) ;
END ;


FUNCTION STRFMONTANT (Tot:string; Total : Extended ; Long,Dec : Integer ; symbole : string3 ; Separateur : Boolean) : string ;
Var
    fact : Double ;
    dd,i : integer;
BEGIN
i := pos ('.', Tot);
dd := 1;
if i <> 0 then
begin
     dd := length (copy(Tot, i+1,length(Tot)));
     // cas 48.2200000
     if (dd > 2) and (copy(Tot, i+3,1) = '0') then dd := 2;
end
else
begin
    if Dec = 0 then
    begin
         Fact:=IntPower(10,2) ;
         Total:=Total*Fact ;
         Result := FloatTostr(Total); exit;
    end;
end;

if dd = Dec then begin Result := FloatTostr(Total); exit; end;
Fact:=IntPower(10,dd) ;
Total:=Total*Fact ;
Fact:=IntPower(10,Dec) ;
Total:=Total/Fact ;
Result := FloatTostr(Total);
END ;


function RechercheOptimise (const TobSrc: tob ; const Field: string ; const Valeur: Variant) : TOB ;
var iStart,iMax,iNumChamp,iCount,iPrevCount: integer ;
begin
     result := nil ;
//     if (Pos ('PHC_',Field) > 0) then TobSrc.Detail.Sort (Field) ;
     if Assigned(TobSrc) and (TobSrc.Detail.Count>0) then
        begin
        // récupération du n° de champ
        iNumChamp := TobSrc.Detail[0].GetNumChamp(Field) ;
        if iNumChamp>0 then
           begin
           iStart     := 0 ;
           iPrevCount := -1 ;
           iCOunt     := 0 ;
           iMax       := TobSrc.Detail.Count-1 ;
           if TobSrc.Detail[iMax].GetValeur(iNumChamp) = Valeur then result := TobSrc.Detail[iMax]
           else if TobSrc.Detail[0].GetValeur(iNumChamp) = Valeur then result := TobSrc.Detail[0]
           else while not assigned(result) do
                   begin
                   if (iPrevCount=iCount) then break ;
                   iPrevCount := iCount ;
                   iCount := ((iMax - iStart) shr 1 + iStart) ;
                   if TobSrc.Detail[iCount].GetValeur(iNumChamp) = Valeur then result :=  TobSrc.Detail[iCount]
                   else if TobSrc.Detail[iCount].GetValeur(iNumChamp) < Valeur then iStart := iCount
                   else iMax := iCount ;
                   end ;
           end ;
        end ;
end ;

function  TCallBackDlg.ExisteEnreg(var Code : string; table,field,ParName : string) : boolean;
var
FTable : TADOTable;
Q1     : TQuery;
begin
        Result := FALSE;
        FTable := TADOTable.Create(Application);
        FTable.Connection := DMImport.DBDonnee as TADOConnection;
        FTable.Tablename := ParName;
        Try FTable.Open; except FTable.Close; exit; end;
        if TRIM(Code) = '' then Code := '';
        if FTable.Locate(field, Code, [loCaseInsensitive]) then
            Result := TRUE;
        FTable.Close;
        if Result then exit;
        Q1 := OpenSQLADO ('SELECT ' + field+' from '+ParName+ ' Where '+field+'="'+ TRIM(Code)+'"', DMImport.DbDonnee);
        if not Q1.EOF then  Result := TRUE
        else Result := FALSE;
        Ferme (Q1);
end;

function TCallBackDlg.ControleCorresp (S : string; ii : integer; Scr : TScript): string;
var
X,pp,i       : integer;
tr1,tr2      : string;
Ctrl         : Boolean;
DEFAUT       : string;
begin
    Ctrl := FALSE;
    for X:=0 to Scr.Champ[ii].TableCorr.FEntree.Count-1 do
    begin
         if Scr.Champ[ii].TableCorr.FEntree.Strings[X]= 'DEFAUT' then
         begin
              if length(S) <= length(Scr.Champ[ii].TableCorr.FSortie.Strings[X]) then
                 DEFAUT := Copy(Scr.Champ[ii].TableCorr.FSortie.Strings[X],1,length(S))
              else
                for i:=length(S)+1 to length(S) do DEFAUT:=Scr.Champ[ii].TableCorr.FSortie.Strings[X]+Scr.Champ[ii].ComplCar ;
         end;
         pp := pos(':',Scr.Champ[ii].TableCorr.FEntree.Strings[X]);
         if pp <> 0 then
         begin
             Ctrl := TRUE;
             tr1 := Copy(Scr.Champ[ii].TableCorr.FEntree.Strings[X],0, pp-1);
             tr2 := Copy(Scr.Champ[ii].TableCorr.FEntree.Strings[X],pp+1, length(Scr.Champ[ii].TableCorr.FEntree.Strings[X]));
             if length(S) <= length(tr1) then
             begin
                 tr2:=Copy(tr2,1,length(S)) ;
                 tr1:=Copy(tr1,1,length(S)) ;
             end
             else
             begin
                for i:=length(S)+1 to length(S) do
                begin
                     tr1:=tr1+Scr.Champ[ii].ComplCar ;
                     tr2:=tr2+Scr.Champ[ii].ComplCar ;
                end;
             end;
             if (S >= tr1) and (S <= tr2) then begin
             Result := S; exit;
             end;
         end;
    end;
    if not Ctrl then exit;
    for i:=length(S) downto 1 do
      for X:=0 to Scr.Champ[ii].TableCorr.FEntree.Count-1 do
      begin
           pp := pos(':',Scr.Champ[ii].TableCorr.FEntree.Strings[X]);
           if pp <> 0 then
           begin
               tr1 := Copy(Scr.Champ[ii].TableCorr.FEntree.Strings[X],0, pp-1);
               tr2 := Copy(Scr.Champ[ii].TableCorr.FEntree.Strings[X],pp+1, length(Scr.Champ[ii].TableCorr.FEntree.Strings[X]));
               begin
                 if Copy(tr1, 0, i) = Copy(S, 0, i) then
                 begin
                      Result := Copy(Scr.Champ[ii].TableCorr.FSortie.Strings[X], 1, length(S)) ;
                      exit;
                 end;
               end;
           end;
      end;
      Result := DEFAUT;
end;



function TCallBackDlg.InsertRecord(const FScript:TScript; FBuffer:PChar) : Boolean;
var
	TR                : PTableRec;
	M, N              : Integer;
	S,S1              : String;
        Val1,Val2,Val3    : string;
        TierName,Valcoll  : string;
        DateDebExe        : string;
        typfield          : integer;
        NN,IndiceG,IndiceN: integer;
        DateTri           : string;
        procedure  AddCompteCache(Cpte : string) ;
        var
        FTable     : TADOTable;
        begin
              if Cpte = '' then exit;
              TierName :=FScript.Name+'Tiers';
              Valcoll := '';
              FTable := TADOTable.Create(Application);
              FTable.Connection := DMImport.DBDonnee as TADOConnection;
              FTable.Tablename := TierName;
              Try FTable.Open; except FTable.Close; exit; end;
              if FTable.Locate('Tiers_CodeAuxiliaire', Cpte, [loCaseInsensitive]) then
                   Valcoll := FTable.FieldByName('Tiers_CompteCollectif').asstring
              else if FTable.Locate('Tiers_CodeAuxiliaire', TRIM(Cpte), [loCaseInsensitive]) then
                   Valcoll := FTable.FieldByName('Tiers_CompteCollectif').asstring;

              FTable.Close;
        end;
        Function OkCompteGene (Cpte : string) : Boolean;
        var
        FTable     : TADOTable;
        begin
              Result := FALSE;
              if Cpte = '' then exit;
              TierName := FScript.Name+'Generaux';
              Valcoll := '';
              FTable := TADOTable.Create(Application);
              FTable.Connection := DMImport.DBDonnee as TADOConnection;
              FTable.Tablename := TierName;
              Try FTable.Open; except FTable.Close; exit; end;
              if FTable.Locate('Generaux_compte', Cpte, [loCaseInsensitive]) then
              begin
                   Result := TRUE; FTable.Close; exit;
              end
              else if FTable.Locate('Generaux_compte', TRIM(Cpte), [loCaseInsensitive]) then
              begin
                   Result := TRUE; FTable.Close; exit;
              end;
              FTable.Close;
        end;
	procedure AffecteChamps(Scr : TScript; TR : PTableRec);
	var N     : Integer;
  Dateper   : TDateTime;
	begin
                Dateper := iDate1900;
		with TR^ do
			for N:=0 to FFields.Count-1 do
			begin
				if (FFields[N].FLon <> 0) and (FFields[N].sel and 4 = 0) then
				begin
					 SetString(S, FBuffer+FFields[N].FOffset, FFields[N].FLon);
                                         if (FFields[N].FFieldName = 'Arjour_code') and (FFields[N+1].FFieldName = 'Arjour_journal') then
                                         begin
                                           NN := 1; while (Scr.Champ[N+NN].binterne) do inc(NN);
                                           SetString(S1, FBuffer+FFields[N+NN].FOffset, FFields[N+NN].FLon);
                                          if ExisteEnreg(S1, 'Arjour', FFields[N+NN].FFieldName, FScript.ParName) then exit;
                                         end;
// ***************** cas COMPTA PGI
                                         if (FFields[N].FFieldName = 'Generaux_fixe') and (FFields[N+1].FFieldName = 'Generaux_ident') then
                                         begin
                                           NN := 2; while (Scr.Champ[N+NN].binterne) do inc(NN);
                                           SetString(S1, FBuffer+FFields[N+NN].FOffset, FFields[N+NN].FLon);
                                           if ExisteEnreg(S1, 'Generaux', FFields[N+NN].FFieldName, FScript.ParName) then exit;
                                         end;

                                         if (FFields[N].FFieldName = 'Tiers_Fixe')then
                                         begin
                                           NN := 2; while (Scr.Champ[N+NN].binterne) do inc(NN);
                                           SetString(S1, FBuffer+FFields[N+NN].FOffset, FFields[N+NN].FLon);
                                           if ExisteEnreg(S1, 'Tiers', FFields[N+NN].FFieldName, FScript.ParName) then exit;
                                         end;

                                         if (FFields[N].FFieldName = 'Journaux_fixe') and (FFields[N+1].FFieldName = 'Journaux_ident')
                                         and (FFields[N+2].FFieldName = 'Journaux_code') then
                                         begin
                                           NN := 2; while (Scr.Champ[N+NN].binterne) do inc(NN);
                                           SetString(S1, FBuffer+FFields[N+NN].FOffset, FFields[N+NN].FLon);
                                           if ExisteEnreg(S1, 'Journaux', FFields[N+NN].FFieldName, FScript.ParName) then exit;
                                         end;

                                         // Fiche 10102
                                         if (FFields[N].FFieldName = 'Ecriture_datemouvement') then
                                           DateTri :=  Copy(S, 5, 4) + Copy(S, 3, 2) + Copy(S, 1, 2);
                                         if (DateTri <> '') and (FFields[N].FFieldName = 'Ecriture_datemvttri') then
                                           S := DateTri;

                                         if (FFields[N].FFieldName = 'Ecriture_comptegeneral') then
                                         begin
                                            Val1 := S; IndiceG := N;
                                         end;
                                         if (FFields[N].FFieldName = 'Ecriture_natureligne') then
                                         begin
                                            Val2 := S; IndiceN := N;
                                            if Val2 = '$' then // pour calculer nature et auxiliaire à partir des tables généraux et Tiers
                                            begin
                                                 if OkCompteGene  (Val1) then
                                                 begin  S := ''; Val2 := '' end;
                                            end;
                                         end;
                                         if (FFields[N].FFieldName = 'Ecriture_auxiliaire') then
                                         begin
                                               if (Val2 = '$') then // pour calculer nature et auxiliaire à partir des tables généraux et Tiers
                                               begin
                                                               FFields[IndiceN].value := 'X';
                                                               S := Val1; Val2 := 'X';
                                                               FFields[IndiceG].value := '';
                                               end ;
                                               if (TRIM(S) <> '') and (Val2 = 'X') then // si uniquement cas compte auxiliaire nature à 'X'
                                               begin
                                                     AddCompteCache(S);
                                                     if (Valcoll <> '') then
                                                        FFields[IndiceG].value := Valcoll;
                                               end
                                               else
                                               if (Val1 <> '') and (TRIM(S) <> '') and (TRIM(Val2) = '') then
                                               begin
                                                    S := '';
                                               end;

                                         end;
// ***************** cas PAIE
                                         if (FFields[N].FFieldName = 'Etablissement1_Fixe') and (FFields[N+1].FFieldName = 'Etablissement1_Ident') then
                                         begin
                                           NN := 3; while (Scr.Champ[N+NN].binterne) do inc(NN);
                                           SetString(S1, FBuffer+FFields[N+NN].FOffset, FFields[N+NN].FLon);
                                           if ExisteEnreg(S1, 'Etablissement1', FFields[N+NN].FFieldName, FScript.ParName) then exit;
                                         end;

                                         if (FFields[N].FFieldName = 'Etablissement2_Fixe') and (FFields[N+1].FFieldName = 'Etablissement2_Ident') then
                                         begin
                                           NN := 2; while (Scr.Champ[N+NN].binterne) do inc(NN);
                                           SetString(S1, FBuffer+FFields[N+NN].FOffset, FFields[N+NN].FLon);
                                           if ExisteEnreg(S1, 'Etablissement2', FFields[N+NN].FFieldName, FScript.ParName) then exit;
                                         end;

// ***************** cas IMOII
                                         if (FFields[N].FFieldName = 'Amort_datedeb') then DateDebExe  := S;
                                         if (FFields[N].FFieldName = 'Amort_metho') then Val1 := S;
                                         if (FFields[N].FFieldName = 'Amort_tauamt') then
                                         begin typfield := FFields[N].FTyp; Val2 := S; end;
                                         if (FFields[N].FFieldName = 'Amort_dure') and(FFields[N].FTyp<>0) then
                                         begin
                                              Val3 := S;
                                              if (FFields[N].FTyp= 1) and (TRIM(Val3) <> '') then
                                              S := FloatToStr(Arrondi ((StrToFloat(Val3)), 0)); // ajout me
                                         end;
                            
                                         if (FFields[N].FTyp= 1) and (FFields[N].FbArrondi = TRUE) then
                                         begin
                                               S1 := TRIM(S);
                                               // c'est le caractère 0A
                                               S1 := retire(' ',S1);
                                               S := FloatToStr(arrondi (StrToFloat(S1), FFields[N].FNbDecimal));
                                         end
                                         else
                                         if (FFields[N].FTyp= 1) and (FFields[N].FNbDecimal <> 2) then
                                         begin
                                               S1 := TRIM(S);
                                               // c'est le caractère 0A
                                               S1 := retire(' ',S1);
                                               if (S1 <> '') and  (not Estalpha(S1))  then
                                                  S := StrfMontant(S1, StrToFloat(S1),13,  FFields[N].FNbDecimal,'',False);
                                         end;

// ***************** cas SISCOII
                                         if (S <> '') and(FFields[N].FFieldName = 'Arecr_daterel') then
                                         Dateper := ConvertDate (S);
                                         if (Dateper <> iDate1900) and (FFields[N].FFieldName = 'Arecr_identecr') then
                                         begin
                                              FFields[N].FTyp:= 1;
                                              FFields[N].Fsel:= 16;
                                              S1 := DateToStr(Dateper);
                                              S := Copy(S1,7,4)+Copy(S1,4,2)+Copy(S1,1,2);
                                         end;
                                         if (OkControle) and (Scr.Champ[N].TableCorr <> nil)  and (TRIM(S) <> '') then
                                         begin
                                          S1 := S;
                                          S := ControleCorresp(S1, N, Scr);
                                         end;
 					 FFields[N].Value := S;
				         TR^.FAnnul := false;
				end;
			end;
	end;

	procedure GetFields( S : String; SL : TStringList);
	var P, P1 : PChar;
		S2 : array [0..255] of char;
	begin
		 P := PChar(S);
		 P1 := P;
		 repeat
			 P := StrScan(P1, ';');
			 if P <> nil then
			 begin
				 SL.Add(StrLCopy(S2, P1, P-P1));
				 Inc(P1);
			 end
			 else
				 SL.Add(P1);
		 until P = nil;
	end;

	procedure AffecteChampsCle(TR : PTableRec);
	var N : Integer;
		S : String;
		SL : TStringList;
	begin
		SL := TStringList.Create;
		GetFields(TR.FTable.IndexDefs[0].Fields, SL);
		for N:=0 to SL.Count-1 do
		begin
			SetString(S, FBuffer+TR.FFields[N].FOffset, TR.FFields[N].FLon);
			TR.FFields[N].Value := S;
		end;
		SL.Free;
	end;
begin
	InsertRecord := False;
	try
	with FScript.Tables do
	  for N:=0 to Count-1 do
		  PTableRec(Objects[N]).FTable.Append;

  { parcours de la liste des champs pour les affecter }
	try
		with FScript do
		begin
			for M:=0 to FScript.Tables.Count-1 do
			begin
                             PTableRec(FScript.Tables.Objects[M]).Fannul := true;
                             AffecteChamps(FScript, PTableRec(FScript.Tables.Objects[M]));
			end;
			{ traitement des operations specifiques }
		  { ------------------------------------- }
		end;
	except
		ShowMessage(Exception(ExceptObject).Message);
		Inc(FAnomalies);
		Anomalies := FAnomalies;
	end;

	with FScript.Tables do
	  for N:=0 to Count-1 do
	  begin
		  TR := PTableRec(Objects[N]);
		  try
			 if not TR^.FAnnul then
                         begin
                              TR^.FTable.Post;
                         end
			 else TR^.FTable.Cancel;
		  except
			  if not FScript.Options.ModifEnrExistant then
			  begin // gestion des violations de clé
				  ShowMessage(Exception(ExceptObject).Message);
				  Inc(FAnomalies);
				  Anomalies := FAnomalies;
				  PTableRec(Objects[N]).FTable.Cancel;
			  end
			  else
			  begin
				  // definition d'un tableau de variant de recherche des champs cles
				  PTableRec(Objects[N]).FTable.Cancel;
				  AffecteChampsCle(PTableRec(Objects[N]));
				  PTableRec(Objects[N]).FTable.Edit;
				  AffecteChamps(FScript, PTablerec(Objects[N]));
				  PTableRec(Objects[N]).FTable.Post;
			  end;
		  end;
	  end;
	except
		ShowMessage(Exception(ExceptObject).Message);
	end;
end;



function TCallBackDlg.LaCallBack(Sender:TObject; AEtape:Integer; FBuffer:PChar; bSelected:Boolean;
var OutputList : TOutputList): integer; stdcall;
var
	N, offset          : integer;
	AScript            : TScript;
	Fr                 : TFieldRec;
        NN                 : string;

        procedure MajCorresp;
        var NTable,table : string;
        M,N              : integer;
        ReqSQl           : string;
        Q1               : Tquery;
        Q2               : TADoTable;
        ChampCode        : string;
        Codedepart       : string;
        Coderemplace     : string;
        Champnature      : string;
        nature           : string;
        begin
                AScript := TScript(Sender);
		with AScript do
		begin
			for M:=0 to Champ.Count-1 do
			begin
                             if Champ[M].ListProfile.count = 0 then continue;
                             table := Copy (Champ[M].Name, 0, pos('_',Champ[M].Name)-1);
                             NTable := AScript.Name+ table;
                             if not FileExists(CurrentDonnee+'\'+NTable) then  break;
                             for N:=0 to Champ[M].ListProfile.Count-1 do
                             begin
                                  // selection du profile
                                  Q1 := OpenSQLADO ('Select * from '+DMImport.GzImpCorresp.TableName
                                  + ' Where Tablename="'+table+'" and Profile="' + Champ[M].ListProfile.strings[N]+'"', DMImport.DbDonnee);

                                  try Q1.Open; except Q1.free; continue; end;
                                  while not Q1.EOF do
                                  begin
                                      ChampCode    := Q1.FindField('Champcode').AsString;
                                      Codedepart   := Q1.FindField('Codedepart').AsString;
                                      Coderemplace := Q1.FindField('Codearrive').AsString;
                                      Champnature  := Q1.FindField('Champnature').AsString;
                                      nature       := Q1.FindField('nature').AsString;
                                      If (Coderemplace <> '') and (nature <> '') then
                                      ReqSQl := 'Update ' + NTable+' set '+table+'_'+Champcode+ '="'+Coderemplace +'",'+
                                      table+'_'+Champnature+'="'+ nature+'"'+
                                      ' Where '+table+'_'+Champcode+' like "'+ Codedepart+'"'
                                      else
                                      If Coderemplace <> '' then
                                      ReqSQl := 'Update ' + NTable+' set '+table+'_'+Champcode+ '="'+Coderemplace +'"'+
                                      ' Where '+table+'_'+Champcode+' like "'+ Codedepart+'"'
                                      else
                                      If (nature <> '') then
                                      ReqSQl := 'Update ' + NTable+' set '+table+'_'+Champnature+'="'+ nature+'"'+
                                      ' Where '+table+'_'+Champcode+' like "'+ Codedepart+'"'
                                      else
                                      continue;


                                      Q2 := TADoTable.Create(nil);
                                      Q2.Connection := DMImport.DBDonnee as TADOConnection;
                                      Q2.TableName := table;
                                      try
                                            Q2.Connection.Execute(ReqSQl);
                                      except
                                            Q2.free; break;
                                      end;
                                      Q2.free;
                                      Q1.next;
                                  end;
                                  Q1.free;
                             end;
			end;
		end;
        end;
begin
	Application.ProcessMessages;

	if AEtape = 0 then       // phase d'initailisation
	begin
		CallBackDlg.Count             := 0;
		CallBackDlg.lblLu.Caption     := '0';
		CallBackDlg.lblsel.Caption    := '0';
		CallBackDlg.Update;
		Result                        := 0;
                OkControle                    := FALSE;
	end
	else if AEtape = 2 then  // phase terminale
	begin
// Parcours pour les tables  pour mise à jour des correspondance
   {$IFNDEF CISXPGI}
                MajCorresp;
    {$ENDIF}
		CallBackDlg.Label1.Caption := 'Fermeture des tables ...';
		CallBackDlg.Update;
		Result := 0;
	end
	else
	begin
		AScript := TScript(Sender);
		if CallBackDlg.prevEtape <> AEtape then
		begin
	 //		CallBackDlg.Tick := GetTickCount;
			CallBackDlg.Label1.Caption := 'Traitement du fichier ...';
			offset := 0;
			for N:=0 to AScript.Champ.Count-1 do
			begin
				FR := AScript.Fields[N];
				FR.FOffset := offset;
				FR.FLon := OutputList.Data^[N].Len;
				Inc(offset, OutputList.Data^[N].Len);
                                if AScript.Champ[N].TableCorr <> nil then
                                begin
                                        OkControle := FALSE;
                                        // si controle de correspondance avec les STD
                                        // exemple 6010000;60199999
                                	(* fiche 21493
                                  for X:=0 to AScript.Champ[N].TableCorr.FEntree.Count-1 do
	                                begin
		                                    if pos(':',AScript.Champ[N].TableCorr.FEntree.Strings[X]) <> 0 then
                                        begin
                                                  OkControle := TRUE;
                                                  break;
                                        end;
                                  end;
                                  *)
                                end;
			end;
		end;

		{ parcours de la liste des tables }
		Inc(CallBackDlg.CountLu);
		CallBackDlg.EnrLu := CallBackDlg.CountLu;
 		if bSelected then
		begin
			CallBackDlg.InsertRecord(AScript, FBuffer);
			Inc(CallBackDlg.Count);
			CallBackDlg.EnrSel := CallBackDlg.Count;
            // si occurence enregistrement = 1
            if (AScript.Champ[0].bUnenreg)  then
            begin
             Result := -1;  CallBackDlg.prevEtape := AEtape; exit;
            end;
		{ mise a jour de la boite de dialogue }
        end;

        if CallBackDlg.bAbort then
           Result := -1
        else
           Result := 1;
        inc (InitProgress) ;
        NN := copy(AScript.Champ[0].Name, 1, pos('_', AScript.Champ[0].Name)-1);
        if not (MoveCurProgressForm (NN+' : '+ IntToStr(InitProgress))) then
        begin
             bAbort := TRUE;
             Result := -1;
        end;
	end;
	CallBackDlg.prevEtape := AEtape;
end;

function TCallBackDlg.CallbackScrut(ANumEnr: Longint; AEtape: Integer): integer; stdcall;
begin
  Application.ProcessMessages;
  if not (MoveCurProgressForm (IntToStr(ANumEnr))) then Result := -1
  else
    if CallBackDlg.bAbort then Result := -1
    else Result := 1;
end;


procedure TCallBackDlg.SetEnrLu(AValue:Integer);
begin
	FEnrLu := AValue;
    lblLu.Caption := IntToStr(FEnrLu);
end;

procedure TCallBackDlg.SetEnrSel(AValue:Integer);
begin
	FEnrSel := AValue;
    lblSel.Caption := IntToStr(FEnrSel);
end;

procedure TCallBackDlg.SetAnomalies(AValue:Integer);
begin
	FAnomalies := AValue;
    lblAnomalies.Caption := IntToStr(FAnomalies);
end;


procedure TCallBackDlg.FormHide(Sender: TObject);
begin
	lblLu.Caption := IntToStr(FEnrLu);
	lblSel.Caption := IntToStr(FEnrSel);
	lblAnomalies.Caption := IntToStr(FAnomalies);
	Update;
end;

procedure TCallBackDlg.FormCreate(Sender: TObject);
begin
	FPrevEnrLu := -1;
	FEnrLu := 0;
	FPrevEnrSel := -1;
	FEnrSel := 0;
	FPrevAnomalies := -1;
	FAnomalies := 0;
        bAbort := false;
        CountLu := 0;
        InitMoveProgressForm (nil, 'Execution Script', 'Traitement en cours ...', 1000, TRUE, TRUE) ;
        InitProgress := 0;

end;

procedure TCallBackDlg.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if not bClosing then Action := caNone;
end;

procedure TCallBackDlg.Free;
begin
  bClosing := True;
  FiniMoveProgressForm;
  inherited Free;
end;

procedure TCallBackDlg.BFermeClick(Sender: TObject);
var Retour : Integer;
begin
  Retour := PGIAsk('Voulez-vous abandonner le traitement en cours ?','Abandon');
  bAbort := Retour = mrYes;
end;

procedure ExportFichierText (var F : TextFile; FichS : string);
var
FLec      : TextFile;
Chaine    : string;
begin
   AssignFile(FLec,FichS);
   {$I-} Reset (FLec) ; {$I+}
   While not EOF(FLec) do
   begin
     MoveCur(FALSE) ;
     Readln(FLec, Chaine);
     WriteLn(F, Chaine)
   end;
   CloseFile(FLec);
end;

Function RenseigneDomaine  (Cle : string) :string;
var
{$IFDEF EAGLCLIENT}
QP : TOB;
{$ELSE}
QP : TQuery;
{$ENDIF}
begin
{$IFDEF CISXPGI}
                  QP := OpenSQL ('SELECT CIS_CLE, CIS_DOMAINE FROM CPGZIMPREQ WHERE CIS_CLE="'
                  +Cle+'"', TRUE);
                  if not QP.EOF then
                     Result := QP.FindField('CIS_Domaine').asstring;
                  Ferme (QP);
{$ELSE}
                  QP := OpenSQLADO ('Select DOMAINE from '+DMImport.GzImpReq.TableName + ' WHERE CLE="'
                  +Cle+'"', DMImport.Db);
                  if not QP.EOF then
                     Result := QP.FieldByName('DOMAINE').asstring;
                  Ferme (QP);
{$ENDIF}
end;

function EcritureFichierExport(var F : TextFile; Sortie, Requete,Domaine  : string; var DatedebutExo, DatefinExo : TDateTime; Lien : string=''; Where : string=''; ScriptName : string=''; Ordre : string=''; Lepays : string='FRA'; PlusLien: TOB=nil; Tbp : TOB=nil): string;
var
Q1        : TQuery;
Chaine    : string;
Formatc   : string;
i,ii,Size : integer;
PPos      : integer;
Contenu   : variant;
OrderBy   : string;
Dateper   : TDateTime;
FMois : TDateTime;
AncienPeriode : string;
AncJournal    : string;
TOBsize       : TOB;
FicheSortie   : string;
St,TableTr    : string;
ListeLien     : TStringList;
XWhere,LLi    : string;
{$IFNDEF CISXPGI}
TPar          : TADOTable;
{$ENDIF}
TT,TTL        : TADOTable;
TableName     : string;
ArContenu     : Variant;
OkAna         : Boolean;
TobParam      : TOB;
{$IFDEF CISXPGI}
TA            : TOB;
{$ENDIF}
IndE          : integer;
ChainAajouter : string;
TF            : TOB;
TableLien     : string;
St1           : string;
jj, ij        : integer;
              procedure recupsize (FieldName : string);
              var
              ia : integer;
              begin
                                   Size := 20;
                                   ia := pos('_', FieldName);
                                   if (ia = 0) then ia := length(FieldName)+1;

                                  {$IFDEF CISXPGI}
                                  if TobParam <> nil then
                                  begin
                                      TA := TobParam.FindFirst(['TABLENAME', 'NOM'],[Copy (FieldName,1,ia-1),Copy (FieldName,ia+1,length(FieldName))], False);
                                      if (TA <> nil) then
                                                Size := TA.GetValue ('LONGUEUR');
                                  end;
                                  {$ELSE}
                                  if TPar.Locate('TableName;Nom',VarArrayOf([Copy (FieldName,1,ia-1),Copy (FieldName,ia+1,length(FieldName))]), [loCaseInsensitive]) then
                                            Size := TPar.FieldByName('Longueur').asinteger;
                                  {$ENDIF}
              end;
              procedure AjouterChamp ( Champ : string);
              begin
                                    recupsize (Champ);
                                    Formatc:= '%-'+IntToStr(size)+'.'+ IntToStr(size)+'s';
                                    ChainAajouter := ChainAajouter + Format(Formatc,[Contenu]);
                                    Chaine := Chaine+Format(Formatc,[Contenu]);
              end;
begin
        Q1 := nil;  TOBsize := nil; ListeLien := nil; size := 0; TT := nil;
{$IFNDEF CISXPGI}
        TPar := nil;
{$ENDIF}
        AncienPeriode := ''; AncJournal := '';
        if (Lien <> '') and (PlusLien = nil)then
        begin
             LLi := Copy(Lien,Pos('/',Lien)+1,length(Lien)-2);
             ListeLien := TStringList.Create;
             St := ReadTokenPipe(LLi,',');
             While (St <> '') do
             begin
                  ListeLien.Add(St);
                  St := ReadTokenPipe(LLi,',');
             end;
             ArContenu := VarArrayCreate([0, ListeLien.Count-1], varVariant);
             if not assigned(TT) then
             begin
                 TT := TADOTable.Create(Application);
                 TT.Connection := DMImport.DbDonnee as TADOConnection;
                 TT.Tablename := ScriptName;
                 Try TT.Open; except TT.Free; end;
             end;
        end;
        if (Domaine = '') then Domaine := GetInfoVHCX.Domaine;
        if (Domaine <> '') and (Domaine <> 'S') and (Domaine <> 'O')then
        begin
{$IFDEF CISXPGI}
             if TbP = nil then RendTobparametre  (Domaine , TobParam, 'COMPTA', 'PARAM', Lepays)
             else
             begin
                  TOBParam := TOB.create ('', nil, -1);
                  TobParam.Dupliquer(Tbp, True, true);
             end;
{$ELSE}
             if not assigned (TPar) then
             begin
                 TPar := TADOTable.create(Application);
                 TPar.TableName := DMImport.GzImpPar.TableName;
                 TPar.Connection := DMImport.Db as TADOConnection;
                 Try TPar.Open; except TPar.Free; end;
             end;
{$ENDIF}
        end;
        try
              PPos := pos('_', Requete);
              if PPos <> 0 then TableTr := Copy (Requete,1, PPos-1)
              else TableTr := Requete;
              if Ordre <> '' then
                 OrderBy := ' order by '+ Ordre
              else
                  OrderBy := '';

              Q1 := OpenSQLADO ('Select * from '+ Sortie + Where+ OrderBy, DMImport.DbDonnee);

              Chaine := '';
              SourisSablier ;   ChainAajouter := '';
              While not Q1.EOF do
              begin
                Application.ProcessMessages;
                XWhere := '';
                  for i:=0 to Q1.FieldCount - 1 do
                  begin
                      Contenu := Q1.FindField(Q1.Fields[i].FieldName).Asstring;
                      if Q1.Fields[i].FieldName = 'Entete_dossier' then
                         FicheSortie :=  Contenu;
                      if Q1.Fields[i].FieldName = ' Exercice_code' then
                         FicheSortie :=  FicheSortie+Contenu;

                      if Domaine = 'S' then // SISCO
                      begin
                            if Q1.Fields[i].FieldName = 'Ardossier_datedebut' then DatedebutExo := ConvertDate (Q1.FindField ('Ardossier_datedebut').asstring);
                            if Q1.Fields[i].FieldName = 'Ardossier_datefin' then  DatefinExo := ConvertDate (Q1.FindField ('Ardossier_datefin').asstring);
                            if Q1.Fields[i].FieldName = 'Arecr_identecr' then Contenu :='';
                            if Q1.Fields[i].FieldName = 'Arecr_code' then
                            begin
                                 Dateper := ConvertDate (Q1.FindField ('Arecr_daterel').asstring);
                                 FMois := FinDeMois(Dateper);
                                 if Dateper > FMois then
                                    Dateper := FMois;
                                 if (AncienPeriode <> Copy(Q1.FindField ('Arecr_daterel').asstring,3, 4)) then
                                 begin
                                      FaitSISCOPER(F, Dateper, DatedebutExo) ;
                                      if AncJournal = Q1.FindField ('Arecr_journal').asstring then
                                      begin
                                          WriteLn(F,'J'+AncJournal+'Journal '+ AncJournal) ;
                                          WriteLn(F,'F001') ;
                                      end;
                                 end;
                                 AncienPeriode := Copy(Q1.FindField ('Arecr_daterel').asstring,3, 4);
                                 if AncJournal <> Q1.FindField ('Arecr_journal').asstring then
                                 begin
                                       AncJournal := Q1.FindField ('Arecr_journal').asstring;
                                       WriteLn(F,'J'+AncJournal+'Journal '+ AncJournal) ;
                                       WriteLn(F,'F001') ;
                                  end;
                                  AncJournal := Q1.FindField ('Arecr_journal').asstring;
                            end;
                      end;
                      if Q1.FieldDefs[i].DataType = ftWideString then
                      begin
                           Formatc:= '%-'+IntToStr(Q1.FieldDefs[i].Size)+'.'+ IntToStr(Q1.FieldDefs[i].Size)+'s';
                      end
                      else
                      if Q1.FieldDefs[i].DataType = ftFloat then
                      begin
                                  if (Domaine = 'S') (*OR (Domaine = 'M')*) then Size := 13
                                  else if (Domaine = 'O') or (Domaine = '') then Size := 20
                                  else
                                  recupsize (Q1.Fields[i].FieldName);
                                  Formatc:= '%'+IntToStr(Size)+'.'+ IntToStr(Size)+'s';
                                  if (Domaine = 'S') OR (Domaine = 'M') then
                                      Contenu := ADroite(Contenu,Size,'0');

                      end;
                      Chaine := Chaine+Format(Formatc,[Contenu]);
                      // maj des champ inexistant du dictionnaire V7
                      if (Domaine = 'C') or (Domaine = 'E') then
                      begin
                             if (Q1.Fields[i].FieldName = 'Ecriture_Tlibre4') then
                             begin
                               if ChainAajouter = '' then
                               begin
                                  if Q1.findField ('Ecriture_refgescom') = nil then
                                     AjouterChamp ('Ecriture_refgescom');
                                  if Q1.findField ('Ecriture_typemvt') = nil then
                                     AjouterChamp ('Ecriture_typemvt');
                                  if Q1.findField ('Ecriture_docid') = nil then
                                     AjouterChamp ('Ecriture_docid');
                                  if Q1.findField ('Ecriture_tresosyn') = nil then
                                     AjouterChamp ('Ecriture_tresosyn');
                                  if Q1.findField ('Ecriture_numtraitechq') = nil then
                                     AjouterChamp ('Ecriture_numtraitechq');
                                  if Q1.findField ('Ecriture_numencdeca') = nil then
                                     AjouterChamp ('Ecriture_numencdeca');
                                  if Q1.findField ('Ecriture_valide') = nil then
                                     AjouterChamp ('Ecriture_valide');
                                  if Q1.findField ('Ecriture_cutoffdeb') = nil then
                                     AjouterChamp ('Ecriture_cutoffdeb');
                                  if Q1.findField ('Ecriture_cutofffin') = nil then
                                     AjouterChamp ('Ecriture_cutofffin');
                                  if Q1.findField ('Ecriture_cutoffdate') = nil then
                                     AjouterChamp ('Ecriture_cutoffdate');
                                  if Q1.findField ('Ecriture_cleecr') = nil then
                                     AjouterChamp ('Ecriture_cleecr');
                                  end
                                  else
                                    if ChainAajouter <> '' then Chaine := Chaine + ChainAajouter;
                             end;
                      end;

                      if (Lien <> '') and (PlusLien = nil) then
                      begin
                           PPos := pos('_', Q1.Fields[i].FieldName);
                           if PPos <> 0 then St := Copy (Q1.Fields[i].FieldName,PPos+1, length(Q1.Fields[i].FieldName))
                           else St := Q1.Fields[i].FieldName;

                           IndE :=  ListeLien.Indexof (St);
                           if IndE >= 0  then
                           begin
                                if Q1.FieldDefs[i].DataType = ftWideString then
                                   St := '="'+Contenu+'"'
                                else
                                   St := '='+Contenu;
                                if XWhere = '' then
                                begin
                                   LLi :=Copy(Lien,0, Pos('/',Lien)-1);
                                   XWhere := ' Where '+ LLi+'_' + ListeLien[IndE] + St;
                                   TableName := LLi+'_' + ListeLien[IndE];
                                end
                                else
                                begin
                                    XWhere := XWhere + ' and '+ LLi+'_' + ListeLien[IndE] + St;
                                    TableName := TableName+';'+LLi+'_' + ListeLien[IndE];
                                end;
                                ArContenu[IndE] :=  Contenu;
                           end;
                      end
                      else
                      if (PlusLien <> nil) then
                      begin
                           ij := pos('_', Q1.Fields[i].FieldName);
                           if ij <> 0 then St := Copy (Q1.Fields[i].FieldName,ij+1, length(Q1.Fields[i].FieldName))
                           else St := Q1.Fields[i].FieldName;
                           if Pos (St, PlusLien.detail[0].getvalue ('ListeLien')) <> 0 then
                           for ij := 0 to PlusLien.detail.count-1 do
                           begin
                                  TF := PlusLien.detail[ij];
                                  if Pos (St, TF.getvalue ('ListeLien')) <> 0 then
                                  begin
                                          LLi :=Copy(TF.getValue('ListeLien'),0, Pos('/',TF.getValue('ListeLien'))-1);
                                          if Q1.FieldDefs[i].DataType = ftWideString then
                                          St1 := '="'+Contenu+'"'
                                          else
                                          St1 := '='+Contenu;
                                          for jj := 0 to TF.detail.count-1 do
                                          begin
                                                   if TF.detail[jj].getnomchamp(jj+1000) = UpperCase (St)  then
                                                   begin
                                                        TF.detail[jj].PutValue (UpperCase (St), Contenu);
                                                       if TF.GetValue('XWhere') = '' then
                                                       begin
                                                                   XWhere := ' Where '+ LLi+'_' + St + St1;
                                                                   TableName := LLi+'_' + St;
                                                       end
                                                       else
                                                       begin
                                                                    XWhere := XWhere + ' and '+ LLi+'_' + St + St1;
                                                                    TableName := TableName+';'+LLi+'_' + St;
                                                       end;
                                                   end;
                                           end;
                                           TF.PutValue('XWhere', XWhere);
                                           TF.PutValue('TableName', TableName);
                                  end;
                           end;

                      end;

                  end;
                  writeln(F, Chaine);
                  Chaine := '';
                  if (Lien <> '') and (XWhere <> '') and (PlusLien = nil) then
                  begin
                       LLi :=Copy(Lien,0, Pos('/',Lien)-1); // nom de la branche analytique
                       Try
                             if ListeLien.Count = 1 then
                               OKAna := TT.Locate(TableName, ArContenu[0], [loCaseInsensitive])
                              else
                               OKAna := TT.Locate(TableName, ArContenu, [loCaseInsensitive]);
                              if OkAna then
                                EcritureFichierExport(F, ScriptName, LLi ,Domaine, DatedebutExo, DatefinExo, '', XWhere, '', '', LePays, nil, TobParam);
                       except end;
                  end
                  else
                  if PlusLien <> nil then
                  begin
                                  for ij := 0 to PlusLien.Detail.count-1 do
                                  begin
                                       TF        := PlusLien.Detail[ij];
                                       St        := TF.GetValue('ListeLien');
                                       TableLien := TF.GetValue('TableLien');
                                       TableName := TF.GetValue('TableName');
                                       XWhere    := TF.GetValue ('XWhere');
            ///                            if Assigned(TT) then begin TT.Close; TT.free; end;
                                       TTL := TADOTable.Create(Application);
                                       TTL.Tablename := TableLien;
                                       TTL.Connection := DMImport.DbDonnee as TADOConnection;
                                       Try
                                       TTL.Open;
                                       except
                                       TTL.Close;  TTL.free;
                                       end;
                                       ArContenu := VarArrayCreate([0, TF.detail.count-1], varVariant);
                                       For IndE:= 0 to TF.detail.Count-1 do
                                             ArContenu[IndE] :=  TF.detail[IndE].GetValeur(IndE+1000);
                                       if TF.Detail.Count = 1 then
                                         OKAna := TTL.Locate(TableName, ArContenu[0], [loCaseInsensitive])
                                        else
                                         OKAna := TTL.Locate(TableName, ArContenu, [loCaseInsensitive]);
                                        if OkAna then
                                        begin
                                             Chaine := '';
                                             for jj :=0 to TTL.FieldCount - 1 do
                                             begin
                                                  Contenu := TTL.FindField(TTL.Fields[jj].FieldName).Asstring;
                                                  if TTL.FieldDefs[jj].DataType = ftWideString then
                                                       Formatc:= '%-'+IntToStr(TTL.FieldDefs[jj].Size)+'.'+ IntToStr(TTL.FieldDefs[jj].Size)+'s'
                                                  else
                                                  if TTL.FieldDefs[jj].DataType = ftFloat then
                                                  begin
                                                              recupsize (TTL.Fields[jj].FieldName);
                                                              Formatc:= '%'+IntToStr(Size)+'.'+ IntToStr(Size)+'s';
                                                  end;
                                                  Chaine := Chaine+Format(Formatc,[Contenu]);
                                             end;
                                             writeln(F, Chaine);
                                        end;
                                          //EcritureFichierExport(F, TableLien, St ,Domaine, DatedebutExo, DatefinExo, '', XWhere, '', '', LePays, nil, TobParam);
                                  end;
                                  TTL.close;  TTL.free; Chaine := '';
                                  for ij := 0 to PlusLien.Detail.count-1 do
                                       PlusLien.Detail[ij].PutValue ('XWhere', '');
                   end;
                  Q1.Next;
              end;
        Finally
              Q1.Free;
{$IFDEF CISXPGI}
              if TbP = nil then TobParam.free;
{$ELSE}
              TPar.Free;
{$ENDIF}
              if TOBsize <> nil then TOBsize.free;
              if (Lien <> '') and Assigned(TT) then TT.Free;
              if (Lien <> '') and  (PlusLien = nil) then begin ListeLien.free; end;

              SourisNormale ;
        end;
end;


{$IFDEF CISXPGI}
{$IFNDEF EAGLCLIENT}
function Recupprefixe( tt : string) : string;
var
Q  : Tquery;
begin
              Q := OpenSQL ('SELECT DT_PREFIXE from DETABLES WHERE DT_NOMTABLE="'+tt+'"', TRUE);
              Result := Q.FindField ('DT_PREFIXE').asstring;
              ferme (Q);
end;
Function ReturnReleveMax (TableTr,CpteGeneral : string) : integer;
var
Num, NumReleve : integer;
QQ             : TQuery;
begin
        QQ:=OpenSQL('SELECT max(EE_NUMERO) FROM ' + GetTableDossier('', TableTr) + ' WHERE EE_GENERAL = "'+CpteGeneral+'"',true) ;
        NumReleve := QQ.Fields[0].AsInteger;
        Ferme(QQ) ;
        QQ:=OpenSQL('SELECT max(CEL_NUMRELEVE) FROM ' + GetTableDossier('', 'EEXBQLIG') + ' WHERE CEL_GENERAL = "'+CpteGeneral+'"',true) ;
        Num := maxintvalue([NumReleve,QQ.Fields[0].AsInteger]);
        Ferme(QQ) ;
        if Num > NumReleve then NumReleve := Num;
        inc(NumReleve);
        Result := NumReleve;
end;

Function MaxRelevEtebac(TableTr,NumCpte : string) : integer;
var
QQ        : TQuery;
NumReleve : integer;
begin
    QQ:=OpenSQL('SELECT max(CET_NUMRELEVE) FROM ' + GetTableDossier('', TableTr) + ' WHERE CET_NUMEROCOMPTE = "'+NumCpte+'"',true) ;
    NumReleve := QQ.Fields[0].AsInteger;
    Ferme(QQ) ;
    inc(NumReleve);
    Result := NumReleve;
end;

function ReturnComptebancaire(ETABBQ,GUICHET,Numerodecompte : string) : string;
var
Q : TQuery;
begin
      Q := Opensql ('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_ETABBQ="'+ETABBQ+'" AND '+
      'BQ_GUICHET="'+GUICHET+'" AND BQ_NUMEROCOMPTE="'+Numerodecompte+'"', TRUE);
      if not Q.EOF then
       Result :=  Q.FindField ('BQ_GENERAL').asstring;
      Ferme (Q);
end;

{$ENDIF}

// Export dans la base
function EcritureDanslaBase(Sortie, Requete,Domaine  : string; var TEcr, TReleve : TOB; var CpteGeneral, RefPointage : string; var NumReleve : integer; var TPControle : TSVControle; Where : string=''; Ordre : string=''; IntegratioReleve :Boolean=TRUE; IntegrationEtebac:Boolean=FALSE): string;
var
Q1        : TQuery;
Chaine    : string;
i,ii      : integer;
Contenu   : variant;
OrderBy   : string;
AncienPeriode : string;
AncJournal    : string;
St,TableTr    : string;
XWhere        : string;
TExp          : TOB;
Prefixe       : string;
TA            : TOB;
TypeAlpha     : string;
ContenuDate   : TDateTime;
ContenuFloat  : double;
NumReleveLig  : integer;
DateReleve    : string;
      procedure CreattobReleve;
      var D,D1 : TOB;
      begin
           if (TableTr <> 'EEXBQ') and (TableTr <> 'EEXBQLIG') and (TableTr <> 'CETEBAC') then exit;
           if (TableTr = 'EEXBQ') then
           begin
                            NumReleve := ReturnReleveMax (TableTr, CpteGeneral);
                            TExp.PutValue ('EE_NUMERO', NumReleve);
                            TExp.PutValue ('EE_GENERAL', CpteGeneral);
                            RefPointage := CpteGeneral+'R'+IntToStr(NumReleve);
                            TExp.PutValue ('EE_REFPOINTAGE', RefPointage);

                            D := TOB.Create('$$$', TReleve, -1);
                            D.AddChampSupValeur('CODE', '01');
                            D.AddChampSupValeur('COMPTE', Texp.GetValue ('EE_GENERAL'));
                            D.AddChampSupValeur('RIB', Texp.GetValue ('EE_RIB'));
                            D.AddChampSupValeur('DEBIT', Texp.GetValue ('EE_OLDSOLDEDEB'));
                            D.AddChampSupValeur('CREDIT',Texp.GetValue ('EE_OLDSOLDECRE'));
                            DateReleve := Texp.GetValue ('EE_DATESOLDE');
                            D.AddChampSupValeur('DATERELEV', DateReleve);
                            D.AddChampSupValeur('DEVISE', Texp.GetValue ('EE_DEVISE'));
                            D.AddChampSupValeur('DATEVALEUR', '');
                            D.AddChampSupValeur('LIBELLE', 'Solde Initial ');
                            D1 := TOB.Create('$$$', TReleve, -1);
                            D1.Dupliquer(D, False, True, True);
                            D1.Putvalue('CODE', '07');
                            D1.PutValue('DEBIT', Texp.GetValue ('EE_NEWSOLDEDEB'));
                            D1.PutValue('CREDIT',Texp.GetValue ('EE_NEWSOLDECRE'));
                            D1.PutValue('LIBELLE', 'Solde Final ');

           end;
           if (TableTr = 'CETEBAC') then
           begin
                        NumReleve := MaxRelevEtebac(TableTr,Texp.GetValue ('CET_NUMEROCOMPTE')) ;
                        TExp.PutValue ('CET_NUMRELEVE', NumReleve);
                        TExp.PutValue ('CET_NUMLIGNE', IntToStr(NumReleveLig)) ;
                        inc(NumReleveLig);

                        if CpteGeneral='' then
                         CpteGeneral := ReturnComptebancaire(Texp.GetValue('CET_ETABBQ'),Texp.GetValue('CET_GUICHET'),Texp.GetValue ('CET_NUMEROCOMPTE'));
                        D := TOB.Create('$$$', TReleve, -1);
                        D.AddChampSupValeur('CODE', Texp.GetValue ('CET_TYPELIGNE'));
                        D.AddChampSupValeur('COMPTE', CpteGeneral);
                        D.AddChampSupValeur('RIB', Texp.GetValue ('CET_NUMEROCOMPTE'));
                        D.AddChampSupValeur('DEBIT', Texp.GetValue ('CET_DEBIT'));
                        D.AddChampSupValeur('CREDIT',Texp.GetValue ('CET_CREDIT'));
                        DateReleve := Texp.GetValue ('CET_DATEOPERATION');
                        D.AddChampSupValeur('DATERELEV', DateReleve);
                        D.AddChampSupValeur('DEVISE', Texp.GetValue ('CET_DEVISE'));
                        DateReleve :=   Texp.GetValue ('CET_DATEVALEUR');
                        D.AddChampSupValeur('DATEVALEUR', DateReleve);
                        if Texp.GetValue ('CET_TYPELIGNE') = '01' then
                        D.AddChampSupValeur('LIBELLE', 'Solde Initial ')
                        else
                        if Texp.GetValue ('CET_TYPELIGNE') = '07' then
                           D.AddChampSupValeur('LIBELLE', 'Solde Final ')
                        else
                        D.AddChampSupValeur('LIBELLE', Texp.GetValue ('CET_LIBELLE'));

           end;
           if (TableTr = 'EEXBQLIG')  then
           begin
                        TExp.PutValue ('CEL_NUMRELEVE', NumReleve) ;
                        TExp.PutValue ('CEL_NUMLIGNE', IntToStr(NumReleveLig)) ;
                        inc(NumReleveLig);
                        TExp.PutValue ('CEL_REFPOINTAGE', RefPointage);
                        TExp.PutValue ('CEL_GENERAL', CpteGeneral);

                        D := TOB.Create('$$$', TReleve, -1);
                        D.AddChampSupValeur('CODE', '04');
                        D.AddChampSupValeur('COMPTE', CpteGeneral);
                        D.AddChampSupValeur('RIB', Texp.GetValue ('CEL_RIB'));
                        D.AddChampSupValeur('DEBIT', Texp.GetValue ('CEL_DEBITEURO'));
                        D.AddChampSupValeur('CREDIT',Texp.GetValue ('CEL_CREDITEURO'));
                        DateReleve := Texp.GetValue ('CEL_DATEOPERATION');
                        D.AddChampSupValeur('DATERELEV', DateReleve);
                        D.AddChampSupValeur('DEVISE', '');
                        DateReleve :=Texp.GetValue ('CEL_DATEVALEUR');
                        D.AddChampSupValeur('DATEVALEUR', DateReleve);
                        D.AddChampSupValeur('LIBELLE', Texp.GetValue ('CEL_LIBELLE'));
           end;
      end;
begin
        Q1 := nil;
        AncienPeriode := ''; AncJournal := '';
         if TPControle = nil then
         begin
              TPControle := TSVControle.create;
              if Domaine = 'X' then TPControle.LeMode := 'IMPORT';
              TPControle.ChargeTobParam(Domaine);
              if TPControle.TOBParam = nil then
                PGIInfo ('Erreur chargement Dictionnaire');
         end;
         if TReleve = nil then   TReleve := TOB.Create('$$$', nil, -1);


        if (Domaine = '') then Domaine := GetInfoVHCX.Domaine;
        try
              ii := pos('_', Requete);
              if ii <> 0 then TableTr := Copy (Requete,1,ii-1)
              else TableTr := Requete;
              if Ordre <> '' then
                 OrderBy := ' order by '+ Ordre
              else
                  OrderBy := '';

              Prefixe := Recupprefixe (TableTr);
              if TEcr = nil then TEcr := TOB.Create('JRL', nil, -1);
              Q1 := OpenSQLADO ('Select * from '+ Sortie + Where+ OrderBy,  DMImport.DbDonnee) ;

              Chaine := '';  NumReleveLig := 1;
              SourisSablier ;
              While not Q1.EOF do
              begin
                Application.ProcessMessages;
                XWhere := '';
                if TableTr = 'BANQUECP' then
                begin
                  CpteGeneral := ReturnComptebancaire(Q1.FindField('BANQUECP_ETABBQ').asstring,Q1.FindField('BANQUECP_GUICHET').asstring,Q1.FindField('BANQUECP_NUMEROCOMPTE').asstring);
                  if cpteGeneral = '' then
                   CpteGeneral :=   BourreOuTronque(Q1.FindField ('BANQUECP_GENERAL').asstring, Fbgene);
                  break;
                end;
                if (not IntegratioReleve) and ((TableTr = 'EEXBQ') or (TableTr = 'EEXBQLIG')) then begin Q1.next; continue; end;
                if (not IntegrationEtebac) and (TableTr = 'CETEBAC')then begin Q1.next; continue; end;
                TExp := TOB.create (TableTr, TEcr, -1);
                  for i:=0 to Q1.FieldCount - 1 do
                  begin
                      ii := pos('_', Q1.Fields[i].FieldName);
                      if ii <> 0 then St := Copy (Q1.Fields[i].FieldName,ii+1, length(Q1.Fields[i].FieldName))
                      else St := Q1.Fields[i].FieldName;
                      if TPControle.TobParam <> nil then
                      begin
                          TA := TPControle.TobParam.FindFirst(['TableName', 'Nom'],[TableTr, St], False);
                          if (TA <> nil) then
                                    TypeAlpha := TA.GetValue ('TypeAlpha');
                      end;

                      if (TypeAlpha = 'A') or (TypeAlpha = '') then
                      begin
                          Contenu := Q1.FindField(Q1.Fields[i].FieldName).Asstring;
                          TExp.AddChampSupValeur (Prefixe+'_'+St, Contenu);
                      end;
                      if TypeAlpha = 'F' then
                      begin
                          if Q1.FindField(Q1.Fields[i].FieldName).asstring ='' then
                             ContenuFloat := 0
                          else
                             ContenuFloat := StrToFloat (Q1.FindField(Q1.Fields[i].FieldName).Asstring);
                          TExp.AddChampSupValeur (Prefixe+'_'+St, ContenuFloat);
                      end;
                      if TypeAlpha = 'D' then
                      begin
                           if Q1.FindField(Q1.Fields[i].FieldName).Asstring= '' then ContenuDate := iDate1900
                           else
                           if IsValidDate (Q1.FindField(Q1.Fields[i].FieldName).Asstring) then
                              ContenuDate := StrToDate(Q1.FindField(Q1.Fields[i].FieldName).Asstring)
                           else  ContenuDate := iDate1900;
                           TExp.AddChampSupValeur (Prefixe+'_'+St, ContenuDate);
                      end;

                  end;
                  CreattobReleve;
                  Q1.Next;
              end;
        Finally
              Q1.Free;
              SourisNormale ;
        end;
end;
{$ENDIF}

end.

