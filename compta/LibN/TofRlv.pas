unit TofRlv;

interface
uses  Windows,StdCtrls,Controls,Classes,db,forms,sysutils,dbTables,ComCtrls,Dialogs,
      HCtrls,HEnt1,HMsgBox,UTOF,UTOB, mul, DBGrids, Vierge, filectrl,FE_Main,
      Graphics,Grids,LettUtil,PrintDBG,ExtCtrls,LettAuto,Ent1,ParamSoc,EtbUser;
Type
     {TOF RECUPRELEVE}
     TOF_RecupReleve = Class (TOF)
     private

       TobReleve : TOB;
       GS : THGrid ;
       CBanque : THValComboBox ;
       CodeDevise : Integer ;
       Procedure InitGrid ;
       procedure OnChangeCBanque(Sender: TObject) ;
       procedure OnClickBAjoutRel(Sender: TObject) ;
       procedure OnClickBChercher(Sender: TObject) ;
       procedure OnClickBReception(Sender: TObject) ;
       procedure OnDblClickG(Sender: TObject) ;
       procedure RempliGrid(RepReleve,CompteBanque  : string);
       procedure RempliSoldeInitial(TR:TOB;StReleve,General,RefPointage : String;NumReleve:integer) ;
       procedure RempliLigne(TRLigne:TOB;StReleve,General,RefPointage : String;NumReleve,NumLigne:integer);
       procedure RempliSoldeFinal(TR:TOB;StReleve:string;NumLigne:integer);
       procedure RempliLibelleComplementaire(TR:TOB;StReleve:string) ;
       procedure RempliMontant(LaTob : TOB; Montant : Double; NomCredit,NomDebit : string) ;
       function VerifiReleve(StReleve,General:string) : Boolean;
       function PutDevise(Devise : string) : integer ;
     public
       procedure OnLoad ; override ;
       procedure OnUpdate ; override ;
       procedure OnIntegreReleve(Nomreleve : string;Efface:Boolean);
     END ;
     {TOF VISURELEVE}
     TOF_VisuReleve = class (TOF)
     private
       FAlternateColor : TColor ;
       GS : THGrid ;
       procedure BImprimerClick(Sender: TObject);
       procedure GetCellCanvas(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
     public
       procedure OnLoad ; override ;
     END ;
     {TOF POINTERELEVE}
     TOF_PointeReleve = class (TOF)
     private
       C_DEBIT,C_CREDIT,RefReleve : string ;
       TobReleve,TobEcriture : TOB;
       MttC,MttB : Double;
       NbCoche : integer ;
       PointageAuto,PointageManu : boolean;
       cBanque,cRefReleve : THValCombobox ;
       GC,GB : THGrid ;
       Rdevise : TRadioGroup ;
       TTotalC,TTotalB : THEdit ;
       procedure OnDblClickGCompta(Sender: TObject) ;
       procedure OnDblClickGBanque(Sender: TObject) ;
       procedure OnChangeCBanque(Sender: TObject) ;
       procedure OnChangeCRefReleve(Sender: TObject) ;
       procedure OnClickRDevise(Sender: TObject) ;
       procedure OnClickBPointageAuto(Sender: TObject) ;
       procedure OnClickBChercher(Sender: TObject) ;
       procedure GetCellCanvasB(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
       procedure GetCellCanvasC(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
       procedure AfficheMontant(TTotal : THEdit;Mtt : Double) ;
       procedure CocheDecoche(G:THGrid) ;
       procedure PointeManu(Sender: TObject) ;       
       procedure ChangeAffichageGrid(ARow : LongInt ; GS : THGrid ) ;
       procedure PositionneDevise ;
       procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState) ;
       procedure GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer) ;
       procedure LesVisibles(Visible : Boolean) ;
       procedure RempliCRefReleve(Banque : string) ;
       procedure RempliCompta ;
       procedure RempliReleve ;
       procedure InitVariables ;
       procedure InitGrid(GS : THGrid) ;       
       function  ConstruitListe(var LM : T_D;var LP : T_I) : integer;
     public
       procedure OnLoad ; override ;
       procedure OnCancel ; override ;
       procedure OnUpdate ; override ;
       procedure OnClose ; override ;
     END;

const   {Grid Releve}
        SR_NOMFICHIER  = 0 ;
        SR_RIB         = SR_NOMFICHIER + 1 ;
        SR_DU          = SR_RIB + 1 ;
        SR_AU          = SR_DU + 1 ;
        SR_DEVISEINI   = SR_AU + 1 ;
        {Grid Visualisation}
        SV_DATE        = 0 ;
        SV_DATEVAL     = SV_DATE + 1 ;
        SV_LIBELLE     = SV_DATEVAL + 1 ;
        SV_DEBIT       = SV_LIBELLE + 1 ;
        SV_CREDIT      = SV_DEBIT + 1 ;
        SV_DEVISE      = SV_CREDIT + 1 ;
        SV_TYPE        = SV_DEVISE + 1 ;
        {Grid Pointage}
        SP_DATE        = 0 ;
        SP_REFERENCE   = SP_DATE + 1 ;
        SP_SOLDE       = SP_REFERENCE + 1 ;
        SP_LIBELLE     = SP_SOLDE + 1 ;
        SP_DATEVAL     = SP_LIBELLE + 1 ;
        SP_POINTER     = SP_DATEVAL + 1 ;
        {Monnaie de Pointage}
        P_FRANC        = 0 ;
        P_EURO         = 1 ;
        P_DEVISE       = 3 ;
        {Autres}
        CARPOINTER     = '+' ;

 	    TMsg: array[01..09] of string 	= (
          {01}         '1;Pointage des Relevés bancaires;Vous ne pouvez pas faire de pointage manuel pendant un pointage automatique ;W;O;O;O'
          {02}        ,'1;Pointage des Relevés bancaires;Le relevé existe déjà, voulez-vous l''integrer ?;Q;YN;N;N'
          {03}        ,'1;Pointage des Relevés bancaires;Aucune banque pour ce relevé;W;O;O;O'
          {04}        ,'1;Pointage des Relevés bancaires;Voulez-vous annuler le pointage en cours ?;Q;YN;N;N'
          {05}        ,'1;Pointage des Relevés bancaires;Voulez-vous enregistrer le pointage ?;Q;YN;N;N'
          {06}        ,'1;Pointage des Relevés bancaires;Voulez-vous abandonnez le pointage an cours ?;Q;YN;N;N'
          {07}        ,'1;Pointage des Relevés bancaires;Vous ne pouvez pas lancer de pointage automatique pendant un pointage manuel ;W;O;O;O'
          {08}        ,'1;Pointage des Relevés bancaires;Voulez-vous effacer le relevé ?;W;YN;N;N'
          {09}        ,'1;Intégration des relevés;Le nom du repertoire ETEBAC3 n''est pas valide dans les paramètres société;W;O;O;O'
          );

implementation
{function z_EclatementReleve(Nomfichier,RepDest : string) : Boolean;
var FSource,FDest: Textfile ; StReleve,NomDest : string ;Ouvert : Boolean ;Year,Month,Day,Hour,Min,Sec,Msec,MsecTmp : word;
begin
Ouvert:=FALSE ; Result:=TRUE;
if AnsiLastChar(RepDest)<>PChar('\') then RepDest:=RepDest+'\' ;
AssignFile(FSource,Nomfichier);
Reset(FSource);
Readln(FSource,StReleve) ;
while not EOF(FSource) do
  Begin
  if copy(StReleve,1,2)='01' then
    begin
    if Ouvert then
      begin
      HshowMessage('1;Reception ETEBAC3;Erreur sur le relevé '+Nomfichier+';W;O;O;O','','') ;
      Result:=FALSE ;
      close(FDest) ;
      end;
    DecodeDate(date(),Year, Month, Day);
    DecodeTime(time(),Hour, Min, Sec,Msec) ;
    MSecTmp:=MSec ; while MSecTmp=MSec do DecodeTime(time(),Hour, Min, Sec,Msec) ;
    NomDest:=RepDest+'Rb'+IntToStr(Month)+IntToStr(Day)+IntToStr(Hour)+IntToStr(Min)+IntToStr(Sec)+IntToStr(MSec)+'.dat' ;
    AssignFile(FDest,NomDest) ;
    ReWrite(FDest) ;
    Ouvert:=TRUE ;
    end ;
  Writeln(FDest,StReleve);
  if copy(StReleve,1,2)='07' then begin close(Fdest) ; Ouvert:=FALSE ; end ;
  Readln(FSource,StReleve) ;
  end ;
if Ouvert then
  begin
  if copy(StReleve,1,2)<>'07' then
    begin
    HshowMessage('1;Reception ETEBAC3;Erreur sur le relevé '+Nomfichier+';W;O;O;O','','');
    Result:=FALSE ;
    end else
    Writeln(FDest,StReleve) ;
  close(FDest) ;
  end ;
close(FSource) ;
end;   }

function ConvertDate(DateChar : string) : TDateTime ;
var Year,Month,Day : word ;
begin
Year  := StrToInt(copy(DateChar,5,2));
if ( Year > 90 )  then Year := Year + 1900 else Year := Year + 2000;
Month:=StrToInt(copy(DateChar,3,2));
Day:=StrToInt(copy(DateChar,1,2));
Result:= EncodeDate(Year,Month,Day);
end;

function ConvertMontant(MontantChar : string;Deci:integer) : Double ;
var MontantFloat : Double;Lettre : char;divi,i : integer ;
begin
Lettre:=MontantChar[14] ;
divi:=1;  MontantFloat :=  1;
for i:=1 to deci do divi := divi * 10;
case Lettre of '{' :       begin Lettre := '0';   end;
               '}' :       begin MontantFloat := -1; Lettre := '0';   end;
               'A'..'I' :  begin Dec(Lettre, 16); end;
               'J'..'R' :  begin MontantFloat := -1; Dec(Lettre, 25); end;
               end;
Result:=MontantFloat*StrToFloat(copy(MontantChar,1,13)+Lettre)/divi;
end;


function TrouveArgument(Argument: String;TypeArg : string): string;
var StArgument : string; i,lg : integer;
begin
lg:=Length(TypeArg)-1 ; StArgument := Argument ; i:=Pos(TypeArg,StArgument) ;
if i>0 then
  begin
  system.Delete(StArgument,1,i+lg) ;
  Result:=ReadTokenSt(StArgument);
  end
  else
  Result:='';
end;

{////////////////////////////////////////////}
{                 RecupReleve                }
{////////////////////////////////////////////}
procedure TOF_Recupreleve.OnLoad;
var BAjoutRel,BChercher,BReception : TButton ;
begin
inherited ;
BAjoutRel:=TButton(GetControl('BAJOUTREL')) ;
if (BAjoutRel <> nil ) and (not Assigned(BAjoutRel.OnClick)) then BAjoutRel.OnClick:=OnClickBAjoutRel;
BChercher:=TButton(GetControl('BCHERCHER')) ;
if (BChercher <> nil ) and (not Assigned(BChercher.OnClick)) then BChercher.OnClick:=OnClickBChercher;
BReception:=TButton(GetControl('BRECEPTION')) ;
if (BReception <> nil ) and (not Assigned(BReception.OnClick)) then BReception.OnClick:=OnClickBReception;
GS:=THGrid(GetControl('GFILES'));
if (GS <> nil) and (not Assigned(GS.OnDblClick)) then GS.OnDblClick:=OnDblClickG;
CBanque:=THValCombobox(GetControl('CBANQUE')) ;
if (CBanque<>nil) and (not Assigned(CBanque.OnChange)) then CBanque.OnClick:=OnChangeCBanque;
RemplirValCombo('TTBANQUECP','','',CBanque.Items,CBanque.Values,False,False) ;
CBanque.Items.Insert(0,'<<Tous>>') ;CBanque.Values.Insert(0,'tous') ;
CBanque.ItemIndex:=0;
//OnChangeCBanque(CBanque);
end;

procedure TOF_Recupreleve.OnUpdate ;
var i : integer ;
Begin
for i := 1 to GS.RowCount-1 do if (GS.Cells[SR_NOMFICHIER,i]<>'')then OnIntegreReleve(GS.Cells[SR_NOMFICHIER,i],TRUE) ;
InitGrid ;
end;

procedure TOF_Recupreleve.InitGrid ;
var i : integer ;
begin
GS.RowCount:=2;GS.FixedRows:=1 ;
GS.FColAligns[SR_RIB]:=taCenter ;
GS.FColAligns[SR_DU]:=taCenter ;
GS.FColAligns[SR_AU]:=taCenter ;
GS.FColAligns[SR_DEVISEINI]:=taCenter ;
for i:=0 to GS.ColCount-1 do GS.Cells[i,1]:= '' ;
end;

procedure TOF_Recupreleve.RempliGrid(RepReleve,CompteBanque : string);
var MySearchRec: TSearchRec; FSource: Textfile ; StReleve,NomReleve : string; QQ : TQuery ;Ok:Boolean;
SQL : string ;
Begin
InitGrid ;
if FindFirst(RepReleve, faAnyFile, MySearchRec) = 0 then
  Begin
  repeat
    NomReleve := ExtractFilePath(RepReleve)+MySearchRec.Name ;
    AssignFile(FSource,NomReleve);
    Reset(FSource);	{ Taille d'enregistrement = 1 }
    if Not EOF(FSource) then
      Begin
      Readln(FSource,StReleve) ;
      Ok:=True;
      if (CompteBanque<>'tous') then
        begin
        SQL:='SELECT BQ_GENERAL FROM BANQUECP WHERE '
                 +'BQ_ETABBQ = "'+copy(StReleve,3,5)+'" AND'
                 +' BQ_GUICHET  = "'+copy(StReleve,12,5)+'" AND'
                 +' BQ_NUMEROCOMPTE = "'+copy(StReleve,22,11)+'" AND'
                 +' BQ_CODE = "'+CompteBanque+'"';
        QQ:=OpenSql(SQL,True);
        if QQ.eof then Ok:=False;
        ferme(QQ);
        end;
      if Ok then
        begin
        GS.Cells[SR_NOMFICHIER,GS.RowCount-1]:= NomReleve;
        GS.Cells[SR_RIB,GS.RowCount-1]:= copy(StReleve,3,5)+' '+copy(StReleve,12,5)+' '+copy(StReleve,22,11);
        GS.Cells[SR_DU,GS.RowCount-1]:=DateToStr(ConvertDate(copy(StReleve,35,6))) ;{Date}
        GS.Cells[SR_DEVISEINI,GS.RowCount-1]:=copy(StReleve,17,3);
        while (copy(StReleve,1,2)<>'07') and (Not Eof(FSource)) do
          Readln(FSource,StReleve) ;
        if copy(StReleve,1,2)='07' then
          GS.Cells[SR_AU,GS.RowCount-1]:=DateToStr(ConvertDate(copy(StReleve,35,6)))
        else
          GS.Cells[SR_AU,GS.RowCount-1]:='  /  /    ';
        GS.RowCount:=GS.RowCount+1 ;
        end;
      end;
    CloseFile(FSource);
   until FindNext(MySearchRec) <> 0 ;
   if ( GS.RowCount <> 2 ) then GS.RowCount := GS.RowCount - 1;
   end;
FindClose(MySearchRec) ;
end;

function TOF_Recupreleve.VerifiReleve(StReleve,General:string) : Boolean;
var QQ:Tquery ;SQL:string ;
begin
Result:=True ;
SQL:='SELECT EE_DATEOLDSOLDE FROM EEXBQ '+
            'WHERE EE_DATEOLDSOLDE = "'+USDateTime(ConvertDate(copy(StReleve,35,6)))+'"'+
            ' AND EE_GENERAL = "'+General+'"';
QQ:=OpenSql(SQL,True);
if Not QQ.Eof then Result:=False ;
ferme(QQ) ;
end;

procedure TOF_RecupReleve.OnChangeCBanque(Sender: TObject) ;
var RepReleve : string ;
Begin
RepReleve:=GetParamSocSecur('SO_REPETEBAC3','');
if AnsiLastChar(RepReleve)<>PChar('\') then RepReleve:=RepReleve+'\' ;
if not DirectoryExists(RepReleve) then HshowMessage(TMsg[9],'','')
                                  else if GS<>nil then RempliGrid(RepReleve+'*.dat',CBanque.Values[CBanque.ItemIndex]);
end;

procedure TOF_Recupreleve.OnClickBChercher(Sender: TObject) ;
begin
OnChangeCBanque(CBanque);
end;

procedure TOF_Recupreleve.OnClickBReception(Sender: TObject) ;
begin
AglLanceFiche('CP','RLVRECEPTION','','','');
end;

procedure TOF_Recupreleve.OnClickBAjoutRel(Sender: TObject) ;
var i : integer ;OpenDialog : TOpenDialog ;
Begin
OpenDialog:=TOpenDialog.create(Ecran);
OpenDialog.InitialDir:='C:\';
OpenDialog.Filter:=TraduireMemoire('Relevés bancaires (*.dat)')+'|*.dat|'+TraduireMemoire('Tous les fichiers (*.*)')+'|*.*';
if OpenDialog.execute then
  begin
  for i:=0 to OpenDialog.Files.Count-1 do
  begin
    if HShowMessage('1;Intégration des relevés;Voulez-vous integrer le relevé '+OpenDialog.Files.Strings[i]+'?;Q;YN;N;N','','')= mrYes then
      //OnIntegreReleve(OpenDialog.Files.Strings[i],FALSE);
      z_EclatementReleve(OpenDialog.Files.Strings[i],GetParamSocSecur('SO_REPETEBAC3','')) ;
    end;
  end;
OpenDialog.free;
OnChangeCBanque(CBanque);
end;

procedure TOF_Recupreleve.OnDblClickG(Sender: TObject) ;
var Argument : string ;
Begin
Argument := 'NOMFICHIER='+GS.Cells[SR_NOMFICHIER,GS.row] ;
AglLanceFiche('CP','RLVVISU','','',Argument);
end;

procedure TOF_RecupReleve.RempliSoldeInitial(TR:TOB;StReleve,General,RefPointage : String;NumReleve:integer) ;
var Montant : double;
begin
TR.PutValue('EE_GENERAL',General);
TR.PutValue('EE_REFPOINTAGE',RefPointage);
TR.PutValue('EE_DATEOLDSOLDE',ConvertDate(copy(StReleve,35,6)) );
TR.Putvalue('EE_DATEPOINTAGE',date()) ;
TR.PutValue('EE_NUMRELEVE',NumReleve );
TR.PutValue('EE_NUMERO',NumReleve );
Montant:=ConvertMontant(copy(StReleve,91,14),StrToInt(copy(StReleve,20,1)) );
RempliMontant(TR,Montant,'EE_OLDSOLDECRE','EE_OLDSOLDECRE') ;
TR.PutValue('EE_RIB',Copy(StReleve,3,5)+Copy(StReleve,12,5)+Copy(StReleve,22,11));
TR.PutValue('EE_DEVISE',copy(StReleve,17,3));
TR.PutValue('EE_ORIGINERELEVE','INT'); //Intégré
TR.PutValue('EE_DATEINTEGRE',date());
TR.PutValue('EE_STATUTRELEVE','NON');
end;

procedure TOF_RecupReleve.RempliLigne(TRLigne:TOB;StReleve,General,RefPointage : String;NumReleve,NumLigne:integer) ;
var Montant : double; Devise : string ;
begin
TRLigne.PutValue('CEL_GENERAL',General);
TRLigne.PutValue('CEL_NUMRELEVE',NumReleve);
TRLigne.PutValue('CEL_NUMLIGNE',NumLigne );
TRLigne.PutValue('CEL_LIBELLE',copy(StReleve,49,31) );
TRLigne.PutValue('CEL_RIB',Copy(StReleve,3,5)+Copy(StReleve,12,5)+Copy(StReleve,22,11)) ;
TRLigne.PutValue('CEL_CODEAFB',copy(StReleve,8,11) );
TRLigne.PutValue('CEL_DATEOPERATION',ConvertDate(copy(StReleve,35,6)));
TRLigne.PutValue('CEL_DATEVALEUR',ConvertDate(copy(StReleve,43,6)));
TRLigne.PutValue('CEL_REFPIECE',copy(StReleve,105,16));
TRLigne.PutValue('CEL_REFPOINTAGE',RefPointage);
TRLigne.PutValue('CEL_IMO',copy(StReleve,21,1));
if copy(StReleve,89,1)<>' ' then TRLigne.PutValue('CEL_EXONERE',copy(StReleve,89,1));
TRLigne.PutValue('NUMECRITURE',copy(StReleve,82,7));
TRLigne.PutValue('NUMLIBELLE','1');
Montant:=ConvertMontant(copy(StReleve,91,14),StrToInt(copy(StReleve,20,1)));
Devise:=copy(StReleve,17,3);
RempliMontant(TRLigne,Montant,'CEL_CREDIT','CEL_DEBIT') ;
end;

procedure TOF_RecupReleve.RempliSoldeFinal(TR:TOB;StReleve:string;NumLigne:integer) ;
var Montant : double;Devise : string ;
begin
TR.PutValue('EE_NBMVT',NumLigne);
TR.PutValue('EE_DATESOLDE',ConvertDate(copy(StReleve,35,6)) );
Montant:=ConvertMontant(copy(StReleve,91,14),StrToInt(copy(StReleve,20,1)));
Devise := copy(StReleve,17,3);
RempliMontant(TR,Montant,'EE_NEWSOLDECRE','EE_NEWSOLDEDEB') ;
end;

procedure TOF_RecupReleve.RempliMontant(LaTob : TOB; Montant : Double; NomCredit,NomDebit : string) ;
var MontantFrf,MontantEur,MontantDev : double ;
begin
MontantFrf:=Montant ; MontantEur:=0 ; MontantDev:=0 ;
case CodeDevise of
1 : MontantFrf:=Montant ;
2 : MontantEur:=Montant ;
3 : MontantDev:=Montant ;
end;

if (Montant>0) then
  begin
  LaTob.PutValue(NomCredit       ,MontantFrf) ;
  LaTob.PutValue(NomCredit+'DEV' ,MontantDev) ;
  LaTob.PutValue(NomCredit+'EURO',MontantEur) ;
  end else
  begin
  LaTob.PutValue(NomDebit       ,MontantFrf*-1) ;
  LaTob.PutValue(NomDebit+'DEV' ,MontantDev*-1) ;
  LaTob.PutValue(NomDebit+'EURO',MontantEur*-1) ;
  end;
end;

procedure TOF_RecupReleve.RempliLibelleComplementaire(TR:TOB;StReleve:string) ;
var NumLibelle : integer; TCh : TOB;
begin
TCh := TR.FindFirst(['NUMECRITURE'],[copy(StReleve,82,7)],TRUE);
if  Tch <> nil  then
  begin
  NumLibelle := StrToInt(TCh.GetValue('NUMLIBELLE'));
  if ( NumLibelle = 1 ) then Tch.PutValue('CEL_LIBELLE1',copy(StReleve,41,41));
  if ( NumLibelle = 2 ) then Tch.PutValue('CEL_LIBELLE2',copy(StReleve,41,41));
  if ( NumLibelle = 3 ) then Tch.PutValue('CEL_LIBELLE3',copy(StReleve,41,41));
  if (NumLibelle = 1) or (NumLibelle = 2) or (NumLibelle = 3) then NumLibelle := NumLibelle + 1;
  Tch.PutValue('NUMLIBELLE',IntToStr(NumLibelle));
  end;
end;

procedure TOF_Recupreleve.OnIntegreReleve(NomReleve : string;Efface:Boolean);
var
  FSource: Textfile ; StReleve,Rib,CodeOpe,Banque,Guichet,Compte,General,RefPointage{,DestFile,DestRepe} : string ;
  TR,TRLigne : TOB ; QQ,Q1 : TQuery ; NumReleve,NumLigne : integer; Integre:Boolean;
begin
AssignFile(FSource,NomReleve);
Reset(FSource);	{ Taille d'enregistrement = 1 }
TobReleve := TOB.create('_RELEVE',nil,-1) ;
TR := nil ; TRLigne := nil ; NumLigne := 0 ; NumReleve := 0 ;
While Not EOF(FSource) do
  Begin
  Readln(FSource,StReleve) ;
  CodeOpe := copy(StReleve,1,2) ;
  //  Gestion de la ligne de solde Entete : Ancien Solde
  if ( CodeOpe = '01' )  then
    Begin
    Banque := copy(StReleve,3,5) ;
    Guichet := copy(StReleve,12,5) ;
    Compte := copy(StReleve,22,11) ;
    Rib := Banque+Guichet+Compte;
    Integre:=True ;
    QQ :=OpenSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE'
                +' BQ_ETABBQ = "'+Banque+'" AND'
                +' BQ_GUICHET  = "'+Guichet+'" AND'
                +' BQ_NUMEROCOMPTE = "'+Compte+'"',True);
    if not QQ.eof then
      begin
      General := QQ.Findfield('BQ_GENERAL').AsString ;
      Q1 :=OpenSQL('SELECT max(EE_NUMERO) as S1 FROM EEXBQ WHERE EE_GENERAL = "'+General+'"',true) ;
      if not Q1.eof then NumReleve := Q1.Findfield('S1').AsInteger + 1 else NumReleve := 1;
      ferme(Q1);
      RefPointage := General+TraduireMemoire('Releve')+IntToStr(NumReleve) ;

      if Not VerifiReleve(StReleve,General) then
        if HShowMessage(TMsg[2],'','') = mrNo then
           Integre:=False;
      if Integre then
        begin
        CodeDevise:=PutDevise(copy(StReleve,17,3));
        TR := TOB.Create('EEXBQ',TobReleve,-1);
        RempliSoldeInitial(TR,StReleve,General,RefPointage,NumReleve)
        end;
      end
    else
      HShowMessage(TMsg[3],'','');
    ferme(QQ);
    end;
  // Gestion des lignes d'écritures
  if (CodeOpe = '04') and (TR <> nil)  then
    Begin
    TRLigne := TOB.Create('EEXBQLIG',TR,-1);
    TRLigne.AddChampSup('NUMECRITURE',False);
    TRLigne.AddChampSup('NUMLIBELLE',False);
    Inc(NumLigne);
    RempliLigne(TRLigne,StReleve,General,RefPointage,NumReleve,Numligne);
    end;
  //Gestion des lignes de commentaires supplementaires
  if (CodeOpe = '05') and (TRLigne <> nil)  then
    RempliLibelleComplementaire(TR,StReleve);
  //Gestion de la ligne de Nouveau Solde
  if (CodeOpe = '07') and (TR <>nil)  then
    RempliSoldeFinal(TR,StReleve,NumLigne) ;
  end ;
  CloseFile(FSource);
  TobReleve.InsertDB(nil,True) ;
  TobReleve.free;
  if Not Efface then if HShowMessage(TMsg[8],'','')= mrYes then Efface:=TRUE ;
  if Efface then DeleteFile(NomReleve) ;
  {DestRepe:=ExtractFilePath(NomReleve)+TraduireMemoire('Archives') ;
  if not DirectoryExists(DestRepe) then CreateDir(DestRepe) ;
  DestFile:=DestRepe+'\'+ExtractFileName(NomReleve);
  RenameFile(NomReleve,DestFile);}
end;

function TOF_Recupreleve.PutDevise(Devise : string) : integer ;
var QQ : TQuery ;
begin
if VH^.TenueEuro = False then
  begin
  if Devise='EUR' then Result:=2
  else
    begin
    QQ:=OpenSQL('SELECT * FROM DEVISE WHERE D_CODEISO="'+Devise+'"',TRUE) ;
    if Not QQ.Eof then
      if V_PGI.DevisePivot=QQ.FindField('D_DEVISE').AsString then Result:=1 else Result:=3
    else
      Result:=-1;
    ferme(QQ) ;
    end ;
  end else
  begin
  if Devise='EUR' then Result:=1
  else
    begin
    QQ:=OpenSQL('SELECT * FROM DEVISE WHERE D_CODEISO="'+Devise+'"',TRUE) ;
    if Not QQ.Eof then
      if QQ.FindField('D_FONGIBLE').AsString='X' then Result:=2 else Result:=3
    else
      Result:=-1;
    ferme(QQ) ;
    end ;
  end ;
end;

{////////////////////////////////////////////}
{                 VisuReleve                 }
{////////////////////////////////////////////}
procedure TOF_VisuReleve.OnLoad;
var FF : TFVierge ;NomFichier,StReleve,CodeOpe  : string;FSource : Textfile ;
    Montant : double ;Deci : integer; BImprimer : TButton;
begin
inherited ;
FF := TFVierge(Ecran);
if (FF<>Nil) then
begin
  if V_PGI.NumAltCol=0 then FAlternateColor:=clInfoBk else FAlternateColor:=AltColors[V_PGI.NumAltCol] ;
  BImprimer:=TButton(GetControl('BIMPRIME')) ;
  if (BImprimer <> nil ) and (not Assigned(BImprimer.OnClick)) then BImprimer.OnClick:=BImprimerClick;
  GS:=THGrid(GetControl('GRLV'));
  if (GS<>nil) then
    begin
    GS.RowCount := 2; GS.ColWidths[SV_TYPE]:=0 ;
    GS.GetCellCanvas:=GetCellCanvas ;
    GS.FColAligns[SV_DEBIT]:=taRightJustify;
    GS.FColAligns[SV_CREDIT]:=taRightJustify;
    NomFichier:=TrouveArgument(FF.FArgument,'NOMFICHIER=') ;
    if ( NomFichier <> '' )  then
      begin
      AssignFile(FSource,NomFichier);
      Reset(FSource);	{ Taille d'enregistrement = 1 }
      While Not EOF(FSource) do
        Begin
        Readln(FSource,StReleve) ;
        CodeOpe := copy(StReleve,1,2) ;
        Deci := StrToInt(copy(StReleve,20,1));
        Montant:=ConvertMontant(copy(StReleve,91,14),Deci);
        if (CodeOpe = '01') or (CodeOpe = '04') or (CodeOpe = '07') then
          Begin
          GS.Cells[SV_DATE,GS.RowCount-1] := DateToStr(ConvertDate(copy(StReleve,35,6))) ;
          if (Montant<0) then GS.Cells[SV_DEBIT,GS.RowCount-1] := FloatToStrF(Montant*-1,ffFixed,20,Deci);
          if (Montant>0) then GS.Cells[SV_CREDIT,GS.RowCount-1] := FloatToStrF(Montant,ffFixed,20,Deci);
          GS.Cells[SV_DEVISE,GS.RowCount-1] := copy(StReleve,17,3);
          if (CodeOpe = '01')  then
            Begin
            FF.Caption := 'Compte : '+copy(StReleve,3,5)+' '+copy(StReleve,12,5)+' '+copy(StReleve,22,11) ;
            GS.Cells[SV_LIBELLE,GS.RowCount-1]:='Solde Initial ';{Libelle}
            GS.Cells[SV_TYPE,GS.RowCount-1] := '1';{}
            end;
          if (CodeOpe = '04') then
            Begin
            GS.Cells[SV_LIBELLE,GS.RowCount-1] := copy(StReleve,49,31);{Libelle}
            GS.Cells[SV_DATEVAL,GS.RowCount-1] := DateToStr(ConvertDate(copy(StReleve,43,6))) ;{Date}
            GS.Cells[SV_TYPE,GS.RowCount-1] := '4';{}
            end;
          if (CodeOpe = '07') then
            begin
            GS.Cells[SV_LIBELLE,GS.RowCount-1] := 'Solde Final';{Libelle}
            GS.Cells[SV_TYPE,GS.RowCount-1] := '7';{}
            end ;
          GS.RowCount := GS.RowCount + 1;
          end;
        end ;
      CloseFile(FSource);
      GS.RowCount := GS.RowCount - 1;
      end;
    end;
  end;
end;

procedure TOF_VisuReleve.BImprimerClick(Sender: TObject);
begin PrintDBGrid(GS,Nil,'','') ; End;

procedure TOF_VisuReleve.GetCellCanvas(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
  if GS.Cells[SV_TYPE,ARow]='1' then
    begin
    GS.Canvas.Font.color:=clBlack;
    GS.Canvas.Brush.Color:=FAlternateColor;
    end ;

  if GS.Cells[SV_TYPE,ARow]='7' then
    begin
    GS.Canvas.Font.Color:=clBlack;
    GS.Canvas.Brush.Color:=FAlternateColor+10 ;
    end ;
//  GS.Invalidate ;
end;


{////////////////////////////////////////////}
{                PointeReleve                }
{////////////////////////////////////////////}
procedure TOF_PointeReleve.OnLoad;
var BChercher,BPointageAuto : TButton ;
begin
inherited ;
GC:=THGrid(GetControl('GCOMPTA'));
if GC<>nil then
  begin
  if not Assigned(GC.OnDblClick) then GC.OnDblClick:=OnDblClickGCompta;
  if not Assigned(GC.OnKeyDown)  then   GC.OnkeyDown:=FormKeyDown;
  if not Assigned(GC.OnMouseUp)  then   GC.OnMouseUp:=GMouseUp;
  GC.FColAligns[SP_SOLDE]:=taRightJustify;
  GC.GetCellCanvas:=GetCellCanvasC ;
  GC.ColWidths[SP_POINTER]:=0;  GC.ColWidths[SP_LIBELLE]:=0; GC.ColWidths[SP_DATEVAL]:=0;
  end;
GB:=THGrid(GetControl('GBANQUE'));
if GB<>nil then
  begin
  if not Assigned(GB.OnDblClick)then  GB.OnDblClick:=OnDblClickGBanque;
  if not Assigned(GB.OnKeyDown)then   GB.OnkeyDown:=FormKeyDown;
  if not Assigned(GB.OnMouseUp)then   GB.OnMouseUp:=GMouseUp;
  GB.FColAligns[SP_SOLDE]:=taRightJustify;
  GB.GetCellCanvas:=GetCellCanvasB ;
  GB.ColWidths[SP_POINTER]:=0;   GB.ColWidths[SP_LIBELLE]:=0;  GB.ColWidths[SP_DATEVAL]:=0;
  end;
RDevise:=TRadioGroup(GetControl('RDEVISE')) ;
if (RDevise<>nil) and (not Assigned(RDevise.OnClick)) then RDevise.OnClick:=OnClickRDevise ;
CBanque:=THValCombobox(GetControl('CBANQUE')) ;
if (CBanque<>nil) and (not Assigned(CBanque.OnChange)) then CBanque.OnChange:=OnChangeCBanque ;
RemplirValCombo('TTBANQUECP','','',cBanque.Items,cBanque.Values,False,False) ;
CBanque.ItemIndex:=0 ;
CRefReleve:=THValCombobox(GetControl('CREFRELEVE')) ;
if (cRefReleve<>nil) and (not Assigned(cRefReleve.OnChange)) then cRefReleve.OnChange:=OnChangecRefReleve ;
BChercher:=TButton(GetControl('BCHERCHER')) ;
if (BChercher<>nil) and (not Assigned(BChercher.OnClick)) then BChercher.OnClick:=OnClickBChercher ;
BPointageAuto:=TButton(GetControl('BPOINTAGEAUTO')) ;
if (BPointageAuto<>nil) and (not Assigned(BPointageAuto.OnClick)) then BPointageAuto.OnClick:=OnClickBPointageAuto ;
TTotalC:=THEDIT(GetControl('TTOTALC')) ;
TTotalC.Font.Style:=TTotalC.Font.Style+[fsBold];
TTotalB:=THEDIT(GetControl('TTOTALB')) ;
TTotalB.Font.Style:=TTotalB.Font.Style+[fsBold];
InitVariables ;
AfficheMontant(TTotalC,0);AfficheMontant(TTotalB,0);
end;

procedure TOF_PointeReleve.LesVisibles(Visible : Boolean) ;
begin
if cBanque<>nil then CBanque.Enabled:=Visible ;
if RDevise<>nil then RDevise.Enabled:=Visible ;
if cRefReleve<>nil then cRefReleve.Enabled:=Visible ;
end;

procedure TOF_PointeReleve.OnCancel ;
var i : integer;
Begin
inherited ;
if HShowMessage(TMsg[4],'','')= mrYes then
  begin
  for i:=1 to GC.RowCount-1 do
    begin
    GC.Cells[SP_POINTER,i]:=' ';
    GC.Cells[SP_REFERENCE,i]:=' ';
    if GC.Cells[SP_DATE,i]<>'' then
      begin
      TobEcriture.Detail[i-1].PutValue('E_REFPOINTAGE','');
      TobEcriture.Detail[i-1].PutValue('E_DATEPOINTAGE',StrToDate('01/01/1900')) ;
      end;
    end;
  GC.Invalidate ;
  for i:=1 to GB.RowCount-1 do
    begin
    GB.Cells[SP_POINTER,i]:=' '; GB.Invalidate ;
    end;
  GC.Invalidate ;
  MttB:=0 ; AfficheMontant(TTotalB,MttB);
  MttC:=0 ; AfficheMontant(TTotalC,MttC);
  NbCoche:=0 ; LesVisibles(TRUE) ;
  PointageAuto:=FALSE;PointageManu:=FALSE;
  end;
end;

procedure TOF_PointeReleve.OnUpdate ;
var i : integer ;
Begin
inherited ;
if (MttB=MttC*-1) and (TobReleve<>nil) and (TobEcriture<>nil) then
  begin
  if HShowMessage(TMsg[5],'','')= mrYes then
    begin
    for i:=1 to GB.Rowcount-1 do
      if (GB.Cells[SP_POINTER,i]=CARPOINTER) and (GB.Cells[SP_DATE,i]<>'') then TobReleve.Detail[i-1].Putvalue('CEL_DATEPOINTAGE',date()) ;
    for i:=1 to GC.Rowcount-1 do
      begin
      if (GC.Cells[SP_POINTER,i]=CARPOINTER) and (GC.Cells[SP_DATE,i]<>'') then
        begin
        TobEcriture.Detail[i-1].Putvalue('E_REFPOINTAGE',RefReleve) ;
        TobEcriture.Detail[i-1].Putvalue('E_DATEPOINTAGE',date()) ;
        end;
      end;
    TobReleve.UpdateDB(TRUE) ;   TobReleve.Free ; TobReleve:=nil ;
    TobEcriture.UpdateDB(TRUE) ; TobEcriture.Free ; TobEcriture:=nil ;
    InitVariables ;
    OnClickBChercher(nil) ;
    LesVisibles(TRUE);
    end;
  end;
end;

procedure TOF_PointeReleve.OnClose ;
Begin
inherited ;
if HShowMessage(TMsg[6],'','')= mrNo then
  Lasterror:=1
else
  begin
  TobReleve.free ; TobReleve:=nil ;
  TobEcriture.free ; TobEcriture:=nil ;
  end;
end;

procedure TOF_PointeReleve.InitVariables ;
begin
C_DEBIT:='' ;
C_CREDIT:='' ;
PointageAuto:=FALSE;
PointageManu:=FALSE;
MttB:=0;
MttC:=0;
if TobReleve<>nil then TobReleve.free; TobReleve:=nil ;
if TobEcriture<>nil then TobEcriture.free ; TobEcriture:=nil ;
NbCoche:=0 ;
RefReleve:='' ;
end;

procedure TOF_PointeReleve.OnChangeCBanque ;
begin
RempliCRefReleve(CBanque.Values[CBanque.ItemIndex]) ;
OnClickBChercher(nil) ;
end;

procedure TOF_PointeReleve.OnChangeCRefReleve ;
begin
RempliReleve ;
end;

procedure TOF_PointeReleve.OnClickRDevise ;
begin
OnClickBChercher(nil) ;
end;

procedure TOF_PointeReleve.PositionneDevise ;
begin
case RDevise.ItemIndex of
  0 : Begin C_DEBIT:='DEBIT' ;     C_CREDIT:='CREDIT';     end;
  1 : Begin C_DEBIT:='DEBITEURO' ; C_CREDIT:='CREDITEURO'; end;
  2 : Begin C_DEBIT:='DEBITDEV' ;  C_CREDIT:='CREDITDEV';  end;
  end;
end ;

procedure TOF_PointeReleve.OnClickBChercher(Sender : TObject) ;
var Okok : boolean ;
Begin
Okok:=TRUE ;
if PointageAuto or PointageManu then if HShowMessage(TMsg[4],'','')= mrNo then Okok:=False ;
if Okok then
  begin
  InitVariables ;  
  PositionneDevise ;
  RempliReleve ;
  RempliCompta ;
  LesVisibles(TRUE) ;
  end;
end;

function  TOF_PointeReleve.ConstruitListe(var LM : T_D;var LP : T_I) : integer ;
var j : integer ;
begin
Result:=0;
if GC.Cells[SP_DATE,1]='' then exit;
FillChar(LM,Sizeof(LM),#0) ; FillChar(LP,Sizeof(LP),#0) ;
for j:=1 to GC.RowCount-1 do
  begin
  if TobEcriture.Detail[j-1].GetValue('E_REFPOINTAGE')='' then
    begin
    LM[j-1]:=TobEcriture.Detail[j-1].GetValue('E_'+C_CREDIT) - TobEcriture.Detail[j-1].GetValue('E_'+C_DEBIT) ;
    LP[j-1]:=0 ;
    end ;
  end ;
Result:=GC.RowCount-1 ;
end;

procedure TOF_PointeReleve.OnClickBPointageAuto(Sender : TObject) ;
Var LM    : T_D ;LP    : T_I ;Solde : double ;Infos : REC_AUTO ;i,j : integer ;
Begin
if PointageManu then begin HShowMessage(TMsg[7],'',''); exit; end;
if GB.Cells[SP_DATE,1]='' then exit ;
PointageAuto:=True ;
Infos.Nival:=0 ; Infos.Decim:=2 ;//DEVI.Decimale ;
Infos.Temps:=0 ; Infos.Unique:=TRUE ;
for i:=1 to GB.RowCount-1 do
begin
Infos.NbD:=ConstruitListe(LM,LP) ;
Solde:=TobReleve.Detail[i-1].GetValue('CEL_'+C_DEBIT) - TobReleve.Detail[i-1].GetValue('CEL_'+C_CREDIT) ;
if LettrageAuto(Solde,LM,LP,Infos)<>0 then
  begin
  GB.Row:=i; CocheDecoche(GB) ;
  for j:=1 to Infos.NbD do
    if LP[j-1]<>0 then
      begin
      TobEcriture.Detail[j-1].PutValue('E_REFPOINTAGE',TobReleve.Detail[i-1].GetValue('CEL_REFPOINTAGE')) ;
      TobEcriture.Detail[j-1].PutValue('E_DATEPOINTAGE',Date());
      GC.Row:=j ; CocheDecoche(GC);
      GC.Cells[SP_REFERENCE,j]:=TobReleve.Detail[i-1].GetValue('CEL_REFPOINTAGE') ;
      end;
  end;
end;
end;

procedure TOF_PointeReleve.ChangeAffichageGrid(ARow : LongInt ; GS : THGrid ) ;
begin
if ARow<=0 then Exit ;
if EstSelect(GS,ARow) then
  begin
  GS.Canvas.Font.Style:=GS.Canvas.Font.Style+[fsItalic];
  GS.Canvas.Brush.Color:=clHighlight;
  end
  else
  begin
  GS.Canvas.Font.Style:=GS.Canvas.Font.Style-[fsItalic] ;
  GS.canvas.Brush.Assign(GS.OldBrush);
  end;
end;

procedure TOF_PointeReleve.GetCellCanvasB(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
var TLibB,TDateVal : THLabel ;
begin
ChangeAffichageGrid(ARow,GB) ;
TLibB:=THLabel(GetControl('TLIBB')) ;
TDateVal:=THLabel(GetControl('TDATEVAL')) ;
TLibB.Caption:=GB.Cells[SP_LIBELLE,GB.Row] ;
TDateVal.Caption:=GB.Cells[SP_DATEVAL,GB.Row] ;
end;

procedure TOF_PointeReleve.GetCellCanvasC(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
var TLibC,TdateEch : THLabel ;
begin
ChangeAffichageGrid(ARow,GC) ;
TLibC:=THLabel(GetControl('TLIBC')) ;
TdateEch:=THLabel(GetControl('TDATEECH')) ;
TLibC.Caption:=GC.Cells[SP_LIBELLE,GC.Row] ;
TdateEch.Caption:=GC.Cells[SP_DATEVAL,GC.Row] ;
end;

procedure TOF_PointeReleve.AfficheMontant(TTotal : THEdit;Mtt : Double) ;
var Mask : string ;
Begin
Mask:=TTotal.EditMask ; if Mask='' then mask:='#,##0.00' ;
if ( Mtt < 0 ) then TTotal.Text:=FormatFloat(mask,0-Mtt)+' D'
               else TTotal.Text:=FormatFloat(mask,Mtt  )+' C';
end;

procedure TOF_PointeReleve.CocheDecoche(G:THGrid) ;
var Solde,Mtt : double ;TR : TOB ;Okok : Boolean ; CeReleve : string ;
begin
if G=GC then
  begin
  Solde:=MttB;
  TR:=TobEcriture.Detail[G.Row-1] ;
  Mtt:=TR.GetValue('E_'+C_CREDIT)-TR.GetValue('E_'+C_DEBIT);
  Okok:=TRUE;
  end
  else
  begin
  Solde:=MttC;
  TR:=TobReleve.Detail[G.Row-1] ;
  Mtt:=TR.GetValue('CEL_'+C_CREDIT)-TR.GetValue('CEL_'+C_DEBIT);
  CeReleve:=TR.GetValue('CEL_REFPOINTAGE');
  if RefReleve='' then RefReleve:=TR.GetValue('CEL_REFPOINTAGE') ;
  if CeReleve=RefReleve then Okok:=TRUE else Okok:=FALSE ;
  end;
  if Okok or PointageAuto then
    begin
    if G.Cells[SP_POINTER,G.Row]=CARPOINTER then
      begin
      G.Cells[SP_POINTER,G.Row]:=' ';
      if G=GC then G.Cells[SP_REFERENCE,G.Row]:=' ';
      Solde:=Solde-Mtt ;
      NbCoche:=NbCoche+1 ;
      end
      else
      begin
      G.Cells[SP_POINTER,G.Row]:=CARPOINTER;
      if G=GC then G.Cells[SP_REFERENCE,G.Row]:=RefReleve;
      Solde:=Solde+Mtt ;
      NbCoche:=NbCoche-1 ;
      end;
    end;
if G=GC then
  begin
  MttB:=Solde ; AfficheMontant(TTotalC,Solde) ;
  end else
  begin
  MttC:=Solde ; AfficheMontant(TTotalB,Solde) ;
  end;
if NbCoche=0 then
  begin
  LesVisibles(TRUE);
  RefReleve:='';
  if Not PointageAuto then PointageManu:=FALSE ;  
  end else
  begin
  LesVisibles(FALSE) ;
  if Not PointageAuto then PointageManu:=TRUE ;
  end ;
G.Invalidate ;
end;


procedure TOF_PointeReleve.PointeManu(Sender: TObject) ;
Begin
if Not PointageAuto then CocheDecoche(THGrid(Sender))
                    else HShowMessage(TMsg[1],'','');
end;

procedure TOF_PointeReleve.OnDblClickGBanque(Sender: TObject) ;
Begin
PointeManu(GB);
end;

procedure TOF_PointeReleve.OnDblClickGCompta(Sender: TObject) ;
Begin
PointeManu(GC) ;
end;

procedure TOF_PointeReleve.InitGrid(GS : THGrid) ;
var i : integer ;
begin
GS.RowCount := 2;GS.FixedRows := 1;
for i:=0 to GS.ColCount-1 do GS.Cells[i,1]:= '';
end;

procedure TOF_PointeReleve.RempliReleve ;
var QQ:TQuery ; SQL : string ;TR : TOB; i :integer;
Begin
if TobReleve<>nil then begin TobReleve.free; TobReleve:=nil; end;
TobReleve:=TOB.create('_RELEVE',nil,-1) ;
SQL:='SELECT * FROM EEXBQLIG WHERE CEL_GENERAL ="'+CBanque.Values[CBanque.ItemIndex]+'"'
           +' AND CEL_DATEPOINTAGE = "'+UsDateTime(StrToDate('01/01/1900'))+'"' ;
if CRefReleve<>nil then
  if CRefReleve.ItemIndex>0 then SQL:=SQL+' AND CEL_REFPOINTAGE = "'+CRefReleve.Items.strings[CRefReleve.ItemIndex]+'"'
                            else SQL:=SQL+' ORDER BY CEL_REFPOINTAGE'
  else SQL:=SQL+' ORDER BY CEL_REFPOINTAGE' ;
QQ:=OpenSQL(SQL,True) ;
TobReleve.LoadDetailDB('EEXBQLIG','','',QQ,FALSE,TRUE) ;
ferme(QQ);
InitGrid(GB) ;
if TobReleve.Detail.Count<>0 then
  begin
  GB.RowCount:=TobReleve.Detail.Count+1;
  for i:=0 to TobReleve.Detail.Count-1 do
    begin
    TR:=TobReleve.Detail[i] ;
    GB.Cells[SP_DATE,i+1]:=TR.GetValue('CEL_DATEVALEUR');
    GB.Cells[SP_REFERENCE,i+1]:=TR.GetValue('CEL_REFPOINTAGE');
    GB.Cells[SP_SOLDE,i+1]:=FormatFloat('#,##0.00',TR.GetValue('CEL_'+C_CREDIT)-TR.GetValue('CEL_'+C_DEBIT)) ;
    GB.Cells[SP_LIBELLE,i+1]:=TR.GetValue('CEL_LIBELLE');
    GB.Cells[SP_DATEVAL,i+1]:=TR.GetValue('CEL_DATEVALEUR');    
    end;
  end;
TobReleve.AddChampSup('POINTER',TRUE);
MttB:=0 ;
AfficheMontant(TTotalB,MttB);
end;

procedure TOF_PointeReleve.RempliCompta ;
var QQ:TQuery ;i : integer ; TR : TOB ;
Begin
if TobEcriture<>nil then begin TobEcriture.free; TobEcriture:=nil; end;
TobEcriture:=TOB.create('_ECRITURE',nil,-1) ;
QQ:=OpenSQL('SELECT * FROM ECRITURE WHERE E_GENERAL ="'+CBanque.values[CBanque.ItemIndex]+'"'
           +' AND E_DATEPOINTAGE = "'+UsDateTime(StrToDate('01/01/1900'))+'"',True) ;
TobEcriture.LoadDetailDB('ECRITURE','','',QQ,FALSE,TRUE) ;
ferme(QQ);
InitGrid(GC) ;
if TobEcriture.Detail.Count<>0 then
  begin
  GC.RowCount:=TobEcriture.Detail.Count+1;
  for i:=0 to TobEcriture.Detail.Count-1 do
    begin
    TR:=TobEcriture.Detail[i] ;
    GC.Cells[SP_DATE,i+1]:=TR.GetValue('E_DATEVALEUR');
    GC.Cells[SP_REFERENCE,i+1]:=TR.GetValue('E_REFPOINTAGE');
    GC.Cells[SP_SOLDE,i+1]:=FormatFloat('#,##0.00',TR.GetValue('E_'+C_CREDIT)-TR.GetValue('E_'+C_DEBIT)) ;
    GC.Cells[SP_LIBELLE,i+1]:=TR.GetValue('E_LIBELLE');
    GC.Cells[SP_DATEVAL,i+1]:=TR.GetValue('E_DATEVALEUR');
    end;
  end;
TobEcriture.AddChampSup('POINTER',TRUE);
MttC:=0 ;
AfficheMontant(TTotalC,0);
end;

procedure TOF_PointeReleve.RempliCRefReleve(Banque : string) ;
var QQ : TQuery ;
Begin
CRefReleve.clear ;
CRefReleve.items.Add('<<Tous>>') ;
QQ:=OpenSQL('SELECT CEL_REFPOINTAGE FROM EEXBQLIG WHERE CEL_GENERAL="'+Banque
           +'" AND CEL_DATEPOINTAGE="'+UsDateTime(StrToDate('01/01/1900'))+'" GROUP BY CEL_REFPOINTAGE',True) ;
while Not QQ.EOF do
  begin
  CRefReleve.items.Add(QQ.FindField('CEL_REFPOINTAGE').AsString) ;
  QQ.Next ;
  end ;
CRefReleve.ItemIndex:=0 ;
if CRefReleve.Items.Count>1 then CRefReleve.Enabled:=TRUE  else CRefReleve.Enabled:=FALSE ;
ferme(QQ) ;
end;

procedure TOF_PointeReleve.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var G:THGrid; OkG,Vide : Boolean ;
begin
G:=THGrid(Sender);
OkG:=G.Focused ; Vide:=(Shift=[]) ;
Case Key of
   VK_SPACE  : if ((OkG) and (Vide)) then PointeManu(G);
   END ;
end;

procedure TOF_PointeReleve.GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
Var G:THGrid;C,R : Longint ;
begin
G:=THGrid(Sender);
if ((ssCtrl in Shift) and (Button=mbLeft)) then
   BEGIN
   G.MouseToCell(X,Y,C,R) ;
   if R>0 then PointeManu(G) ;
   END ;

end;

Initialization
registerclasses([TOF_RecupReleve]);
registerclasses([TOF_VisuReleve]);
registerclasses([TOF_PointeReleve]);
end.
