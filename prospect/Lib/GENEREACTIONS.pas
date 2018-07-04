unit GENEREACTIONS;


interface

uses  StdCtrls,Controls,Classes,forms,sysutils,
      HCtrls,HEnt1,HMsgBox,UTOF,
      UTOM,AGLInit,Hstatus,M3FP,UTOB,Hqry,
{$IFDEF GIGI}
      EntGc, HTB97,
{$ENDIF}
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul,eFichList,
{$ELSE}
      HDB,db,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
      FichList,Fe_Main,mul,
{$ENDIF}
{$IFDEF AFFAIRE}
{$IFNDEF GRCLIGHT}
   UtofMailingaff,
{$ENDIF GRCLIGHT}
      UtofAfTraducChampLibre,
{$ENDIF}
      Vierge,Ed_Tools;

Const
    Encours : string = 'PRE';
    MaxCol  : Integer = 50;

Type
    TOM_ACTIONSGENERIQUES = Class (TOM)
    Private
        codeoperation : string;
        Old_Typeaction,stProduitpgi: string;
        soRtactgestech: boolean;
        soRtactgestcout: boolean;
        Old_DateAction : TDateTime;

    procedure LookEcranActGen (Affect : Boolean);
    procedure ChgtTypeActionGen;
    procedure OnExit_DateActionGen(Sender: TObject);

    Public
        procedure OnNewRecord ; override ;
        procedure OnUpdateRecord ; override ;
        procedure OnLoadRecord ; override ;
        procedure OnChangeField (F : TField) ; override ;
        procedure OnArgument (Arguments : String )  ; override ;
        Procedure OnDeleteRecord ; override ;
     END;

{$ifdef AFFAIRE}
                //mcd 24/11/2005 pour faire affectation depuis ressource si paramétré
    TOF_RtSelectProsp = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
    TOF_RtSelectProsp = Class (TOF)
{$endif}
    Private
            MajContact : Boolean;
    Public
        Stoperation,stProduitpgi : string;
        noactgen:integer;

        procedure OnLoad ; override ;
        procedure GenerationDesActions(POperation : string;PActgen:integer);
//        procedure CreationAction(prospect,codetiers,commercial:string;contact:integer;TobInfosAction:TOB ) ;
        procedure OnArgument (Arguments : String ); override ;
{$ifdef GIGI}
        procedure bChercheClick(Sender: TObject);
{$endif}
     Private
        procedure DoAlimDureeAction(TobActGen : tob);
     END;

    TOF_RtActionsLot = Class (TOF)
    Private
        Old_Typeaction,TypeTrait,StParent,stProduitpgi : string;
        iNumChainage,ErreurLot : integer;
        dureeAct : double;
        SauvBlocNote : string;
        procedure LookEcranActLot (Affect : Boolean);
        procedure ChgtTypeActionLot;
        procedure OnExit_DateActionLot(Sender: TObject);
    Public
        procedure OnArgument (Arguments : String ); override ;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
     END;

    TOF_RtGenereAct_Fic = Class (TOF)
    Private
        TobActGen : Tob;
        ErreurFic : Boolean;
    Public
        procedure OnArgument (Arguments : String ); override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
        procedure AffColTiers (AffAno : Boolean) ;
        procedure AffColBlocNote (AffAno : Boolean);
     END;

procedure AGLGenereAction(parms:array of variant; nb: integer ) ;
procedure GenereAction (operation,argument:string);
procedure AGLGenerationDesActions(parms:array of variant; nb: integer ) ;
Function RTLanceFiche_RtActionsLot(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

const
	// libellés des messages
	TexteMessage: array[1..15] of string 	= (
          {1}        'La date d''échéance doit être postérieure à la date d''action'
          {2}        ,'La date de l''action doit être renseignée'
          {3}        ,'Vous ne pouvez pas effacer cette action générique, elle est actuellement utilisée dans les actions.'
          {4}        ,'Le type d''action doit être renseigné.'
          {5}        ,'Le libellé doit être renseigné.'
          {6}        ,'Le responsable doit être renseigné'
          {7}        ,'Le responsable n''existe pas'
          {8}        ,'Vous n''êtes pas autorisé à créer cette action'
          {9}        ,'L''opération n''existe pas'
          {10}       ,'Génération des actions'
          {11}       ,'Traitement en cours'
          {12}       ,'Impossible de créer l''action pour le tiers %s'
          {13}       ,'par client'
          {14}       ,'par contact'
          {15}       ,'à partir d''un fichier'
          );

implementation
uses UtilSelection,ParamSoc,UtilRT,EntRT,
      UtilGC,UtilArticle,UtilAction,variants;

Type T_InsertAction = Class(TObject)
     Prospect : string;
     CodeTiers : string;
     Commercial : string;
     Contact : integer;
     TobInfosAction : TOB ;
     Procedure InsertAction ;
     End ;

procedure AGLGenereAction(parms:array of variant; nb: integer ) ;
begin
GenereAction (String(parms[0]),String(parms[1]));
end;

procedure AGLGenerationDesActions(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_RtSelectProsp) then TOF_RtSelectProsp(TOTOF).GenerationDesActions('',0)
{$IFDEF AFFAIRE}
{$IFNDEF GRCLIGHT}
else if (TOTOF is TOF_MailingAff) then TOF_RtSelectProsp(TOTOF).GenerationDesActions(Parms[1],Parms[2])
{$ENDIF GRCLIGHT}
{$ENDIF}
else exit;
end;

procedure AGLChgtTypeActionGen( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFicheListe) then OM:=TFFicheListe(F).OM else exit;
if (OM is TOM_ACTIONSGENERIQUES) then TOM_ACTIONSGENERIQUES(OM).ChgtTypeActionGen else exit;
end;

procedure AGLChgtTypeActionLot( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit;
if (TOTOF is TOF_RTActionsLot) then TOF_RTActionsLot(TOTOF).ChgtTypeActionLot else exit;
end;

procedure GenereAction (operation,argument:string);
var act,stArg:string;
begin
if (pos('GRF', argument) <> 0) then stArg := ';PRODUITPGI=GRF'
else stArg := ';PRODUITPGI=GRC';
act:=AGLLanceFiche ('RT','RTINFOS_ACTGEN',operation,'','CODEOPER='+operation+stArg);
if (act <> '') then
   begin
   if (pos('GRF', argument) <> 0) then
     begin
     if (pos('FOU', argument) <> 0) then
        AGLLanceFiche ('RT','RFPROS_MUL_SELECT','OPERATION='+operation,'','CODEOPER='+operation+';ACTGEN='+act+stArg)
     else
        AGLLanceFiche ('RT','RFLIGN_MUL_SELECT','OPERATION='+operation,'','CODEOPER='+operation+';ACTGEN='+act+stArg);
     end
   else
     begin
     if (pos('CLI', argument) <> 0) then
        AGLLanceFiche ('RT','RTPROS_MUL_SELECT','OPERATION='+operation,'','CODEOPER='+operation+';ACTGEN='+act)
  {$IFDEF AFFAIRE}
     else if (pos('AFF', argument) <> 0) then
        AGLLanceFiche ('RT','RTAFF_MAILIN_CONT','OPERATION='+operation,'','CODEOPER='+operation+';ACTGEN='+act)
  {$ENDIF}
     else
     if (pos('LIGN', argument) <> 0) then
        AGLLanceFiche ('RT','RTLIGN_MUL_SELECT','OPERATION='+operation,'','CODEOPER='+operation+';ACTGEN='+act)
     else
        AGLLanceFiche ('RT','RTPRO_SELECT_CONT','OPERATION='+operation,'','CODEOPER='+operation+';ACTGEN='+act);
     end;
   end;
end;

procedure CreationAction(prospect,codetiers,commercial:string;contact:integer;TobInfosAction:TOB);
var InsertAct:T_InsertAction;
begin
InsertAct:=T_InsertAction.Create;
InsertAct.prospect:=prospect;
InsertAct.codetiers := codetiers;
InsertAct.commercial := commercial;
InsertAct.contact := contact;
InsertAct.TobInfosAction:=TobInfosAction ;
if Transactions(InsertAct.InsertAction,3) <> oeOK then PGIBox(format(TraduireMemoire(TexteMessage[12]),[codetiers]), TexteMessage[10]) ;
InsertAct.free;
end;

procedure T_InsertAction.InsertAction;
var Q : TQuery;
    nbact:integer;
    Select,Utilisateur : string;
begin
Select := 'SELECT MAX(RAC_NUMACTION) FROM ACTIONS WHERE RAC_AUXILIAIRE = "'+ prospect+'"';
Q:=OpenSQL(Select, True,-1, '', True);
if not Q.Eof then
   begin
   nbact := Q.Fields[0].AsInteger;
   TobInfosAction.PutValue ('RAC_NUMACTION',nbact+1);
   end
else
   TobInfosAction.PutValue ('RAC_NUMACTION',1);
Ferme(Q) ;

//if (StOperation <> '') and (TobInfosAction.GetValue ('RAC_INTERVENANT') ='') then
if (TobInfosAction.GetValue ('RAC_INTERVENANT') ='') then
    begin
    //TobInfosAction.PutValue ('RAC_INTERVENANT','');
    if commercial <> '' then
        begin
        Select := 'SELECT GCL_UTILASSOCIE FROM COMMERCIAL WHERE GCL_COMMERCIAL = "'+ commercial+'"';
        Q:=OpenSQL(Select, True,-1, '', True);
        if not Q.Eof then
           Utilisateur := Q.Fields[0].AsString;
        Ferme(Q) ;
        if Utilisateur <> '' then
            begin
            Select := 'SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_UTILASSOCIE = "'+ Utilisateur+'"';
            Q:=OpenSQL(Select, True,-1, '', True);
            if not Q.Eof then
                begin
                TobInfosAction.PutValue ('RAC_INTERVENANT',Q.Fields[0].AsString);
{                TobInfosAction.PutValue ('RAC_INTERVINT',Q.Fields[0].AsString+';'); }
                end;
            Ferme(Q) ;
            end;
        end;
     end;
TobInfosAction.PutValue ('RAC_AUXILIAIRE',prospect);
TobInfosAction.PutValue ('RAC_TIERS',codetiers);
TobInfosAction.putvalue ('RAC_DATECREATION', Date);
TobInfosAction.putvalue ('RAC_NUMEROCONTACT', contact);
TobInfosAction.InsertDB(Nil);
if V_PGI.IoError = oeOK then RTCreationLiensActions (TobInfosAction);
end;

procedure TraitementFic (TobAction,TobActGen : Tob;StFichier : string; ColTiers: integer; ColContact : string);
var FLig : textfile;
    StLig,CodeTiers, CodeProspect, commercial, Select,Mess,Critere,StColContact : string;
    Liste,Col,SavCodeTiers : string;
    NoContact,nbActions,i,NbCol,Nbenr : integer;
    lg : double;
    Q : TQuery;
    BlocNote : TStrings;
    Trouve,Trt : boolean;
    Searchrec : TSearchRec;
    ChampsLignes: array[1..50] of string;
begin
Nbenr := 0;
Lg := 0;
AssignFile(FLig, stFichier);
Reset (FLig);
while not EOF(FLig) do
  begin
  readln (FLig,StLig);
  Inc (Nbenr);
  Lg := lg + Length(Stlig);
  if (Nbenr >= 10 ) then break;
  end;
if (Nbenr <> 0) then Lg := Lg / Nbenr;
if (lg <> 0) then
  begin
  FindFirst(stFichier, faAnyFile, SearchRec);
  Nbenr := Trunc(SearchRec.Size / lg);
  end;

Reset (FLig);
nbActions := 0;
SavCodeTiers := '';
Trouve := False;
BlocNote:=TStringList.Create;
NbCol:= Coltiers;
StColContact := ColContact;
Repeat
Col:=trim(ReadTokenSt(StColContact));
if Col<>'' then
  begin
  if Valeuri(Col) > NbCol then
    begin
    NbCol := Valeuri(Col)
    end;
  end;
until  Col='';

//InitMove(Nbenr,'');
InitMoveProgressForm (Nil, TexteMessage[10],TexteMessage[11], Nbenr,false,false) ;
if not EOF(FLig) then
  begin
  readln (FLig,StLig);
//  MoveCur (False);
  MoveCurProgressForm ('');

  Liste := '';
  for i := 1 to MaxCol do
    begin
    ChampsLignes[i] := '';
    end;
  for i := 1 to NbCol do
    begin
    Critere:=trim(ReadTokenSt(StLig));
    if i = ColTiers then CodeTiers := Critere
    else ChampsLignes[i] := Critere;
    end;
  StColContact := ColContact;
  Repeat
  Col:=trim(ReadTokenSt(StColContact));
  if Col<>'' then
    begin
    Liste := Liste + ChampsLignes[Valeuri(Col)] + ' ';
    end;
  until  Col='';

  Trt := True;

  while (Trt = True) do
    begin
    if EOF(FLig) then Trt := False;
    if (SavCodeTiers <> CodeTiers) then
      begin
      if (SavCodeTiers <> '') then
        begin
        if Trouve then
          begin
          TobAction.PutValue ('RAC_BLOCNOTE',BlocNote.Text);
          TobAction.PutValue ('RAC_INTERVENANT',TobActGen.GetValue('RAG_INTERVENANT'));
{          if TobActGen.GetValue('RAG_INTERVENANT')<>'' then
             TobAction.PutValue ('RAC_INTERVINT',TobActGen.GetValue('RAG_INTERVENANT')+';')
          else
             TobAction.PutValue ('RAC_INTERVINT','');  }
          nocontact := 0;
          Select := 'SELECT C_NUMEROCONTACT FROM CONTACT WHERE C_TYPECONTACT = "T" AND C_PRINCIPAL = "X" AND C_FERME <> "X" AND C_AUXILIAIRE = "'+ CodeProspect+'"';
          Q:=OpenSQL(Select, True,-1, '', True);
          if not Q.Eof then
            nocontact := Q.FindField('C_NUMEROCONTACT').AsInteger;
          Ferme(Q) ;
          CreationAction (codeprospect,SavCodetiers,commercial,nocontact,TobAction);
          Inc (NbActions);
          end;
        end;
      CodeProspect := '';
      Commercial := '';
      BlocNote.Clear;
      Trouve := False;
      if (CodeTiers <> '') then
        begin
        Select := 'SELECT T_AUXILIAIRE,T_REPRESENTANT FROM TIERS WHERE T_TIERS = "'+ CodeTiers+'"';
        Q:=OpenSQL(Select, True,-1, '', True);
        if not Q.Eof then
          begin
          CodeProspect := Q.FindField('T_AUXILIAIRE').AsString;
          Commercial := Q.FindField('T_REPRESENTANT').AsString;
          Trouve := True;
          end;
        Ferme(Q) ;
        end;
      end;

    SavCodeTiers := CodeTiers;
    with BlocNote do
      Add(Liste);
    readln (FLig,StLig);
//    MoveCur (False);
    MoveCurProgressForm ('');

    Liste := '';
    for i := 1 to MaxCol do
      begin
      ChampsLignes[i] := '';
      end;
    for i := 1 to NbCol do
      begin
      Critere:=trim(ReadTokenSt(StLig));
      if i = ColTiers then CodeTiers := Critere
      else ChampsLignes[i] := Critere;
      end;
    StColContact := ColContact;
    Repeat
    Col:=trim(ReadTokenSt(StColContact));
    if Col<>'' then
      begin
      Liste := Liste + ChampsLignes[Valeuri(Col)] + ' ';
      end;
    until  Col='';

    end; // fin while

  end;
// la derniére action n'est pas créée
if Trouve then
  begin
  TobAction.PutValue ('RAC_BLOCNOTE',BlocNote.Text);
  TobAction.PutValue ('RAC_INTERVENANT',TobActGen.GetValue('RAG_INTERVENANT'));
{  if TobActGen.GetValue('RAG_INTERVENANT')<>'' then
     TobAction.PutValue ('RAC_INTERVINT',TobActGen.GetValue('RAG_INTERVENANT')+';')
  else
     TobAction.PutValue ('RAC_INTERVINT',''); }
  nocontact := 0;
  Select := 'SELECT C_NUMEROCONTACT FROM CONTACT WHERE C_TYPECONTACT = "T" AND C_PRINCIPAL = "X" AND C_FERME <> "X" AND C_AUXILIAIRE = "'+ CodeProspect+'"';
  Q:=OpenSQL(Select, True,-1, '', True);
  if not Q.Eof then
    nocontact := Q.FindField('C_NUMEROCONTACT').AsInteger;
  Ferme(Q) ;
  CreationAction (codeprospect,SavCodetiers,commercial,nocontact,TobAction);
  Inc (NbActions);
  end;
//FiniMove ;
FiniMoveProgressForm;

CloseFile(FLig);
BlocNote.Free;
Mess:='Nombre d''actions générées : '+IntToStr(NbActions);
PGIBox(Mess,'Informations sur la génération');
end;

procedure TOM_ACTIONSGENERIQUES.OnChangeField(F: TField);
begin

end;

procedure TOM_ACTIONSGENERIQUES.OnLoadRecord;
var i:integer;
    Affect : Boolean;
begin
inherited ;
for i:=1 to 3 do
    SetControlCaption('TRAG_TABLELIBRE'+IntToStr(i),'&'+RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(i),FALSE)) ;
if Getfield ('RAG_TYPEACTION') <> '' then
   begin
   Old_Typeaction := Getfield ('RAG_TYPEACTION');
   Affect := False;
   LookEcranActGen(Affect);
   end;
Old_DateAction := Getfield ('RAG_DATEACTION');
if  not (ds.state in [dsinsert]) then
    begin
    if ExisteSQL('SELECT RAC_LIBELLE FROM ACTIONS WHERE RAC_NUMACTGEN='
     + IntToStr(GetField('RAG_NUMACTGEN'))+' AND RAC_OPERATION="'+codeoperation+'"') then ModifAutorisee := False;
    end;
end;

procedure TOM_ACTIONSGENERIQUES.OnNewRecord;
begin
inherited;
SetField('RAG_DATEECHEANCE', iDate2099);
SetField('RAG_ETATACTION', Encours);
SetField('RAG_PRODUITPGI',stProduitpgi);
SetField('RAG_DATEACTION', V_PGI.DateEntree);
Old_Typeaction := '';
end;

procedure TOM_ACTIONSGENERIQUES.OnUpdateRecord;
var Q : TQuery;
    nbact:integer;
    Select : string;

begin
inherited ;
if GetField ('RAG_OPERATION') <> 'MODELES D''ACTIONS' then
    begin
    if GetField ('RAG_DATEACTION') = iDate1900 then
       begin
       Lasterror:=2;
       LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
       SetFocusControl('RAG_DATEACTION') ;
       exit;
       end;
   end;

if GetField ('RAG_INTERVENANT') <> '' then
    begin
    if (Not ExisteSQL ('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+GetField('RAG_INTERVENANT')+'"')) then
       begin
       SetFocusControl('RAG_INTERVENANT');
       LastError := 7;
       LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
       exit;
       end;
    end;

if soRtactgestech then
   begin
   if GetField ('RAG_DATEECHEANCE') < GetField ('RAG_DATEACTION') then
      begin
      Lasterror:=1;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      SetFocusControl('RAG_DATEECHEANCE') ;
      exit;
      end;
   end;

if getfield ('RAG_NUMACTGEN') = 0 then
    begin
    Select := 'SELECT MAX(RAG_NUMACTGEN) FROM ACTIONSGENERIQUES WHERE RAG_OPERATION = "'+codeoperation+'"';
    Q:=OpenSQL(Select, True,-1, '', True);
    if not Q.Eof then
       begin
       nbact := Q.Fields[0].AsInteger;
       setfield ('RAG_NUMACTGEN',nbact+1);
       end
       else
       setfield ('RAG_NUMACTGEN',1);
    Ferme(Q) ;
    setfield ('RAG_OPERATION',codeoperation);
    end;
     

end;

procedure TOM_ACTIONSGENERIQUES.OnArgument (Arguments : String );
var Critere,ChampMul,ValMul : string;
    x : integer;
begin
inherited ;
  codeoperation:='';
  stproduitpgi := '';

  TCheckBox(GetControl('CONSULTATION')).state:=cbUnChecked;
  Repeat
      Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
      if Critere <> '' then
      begin
          x:=pos('=',Critere);
          if x<>0 then
          begin
             ChampMul:=copy(Critere,1,x-1);
             ValMul:=copy(Critere,x+1,length(Critere));
             if ChampMul='CODEOPER' then
             begin
                codeoperation := ValMul;
             end;
             if ChampMul='PRODUITPGI' then
                stProduitpgi := ValMul;
          end
          else
          begin
          x := pos('OPERATION',Critere);
          if x <> 0 then
             begin
             TCheckBox(GetControl('CONSULTATION')).state:=cbChecked;
             SetControlVisible ('BSELECTACTGEN',False);
             end;
          end;
      end;
  until  Critere='';
if stProduitpgi = '' then stProduitpgi := 'GRC';
if codeoperation = 'MODELES D''ACTIONS' then
     begin
     SetControlEnabled('RAG_DATEACTION',FALSE) ;
     SetControlEnabled('TRAG_DATEACTION',FALSE) ;
     SetControlVisible ('BSELECTACTGEN',False);
     SetControlVisible ('BPROSPECT',False);
     TFFicheListe(Ecran).HelpContext:=111000245;
     TFFicheListe(Ecran).Caption:=TraduireMemoire('Modèles d''actions CTI :');
     end;

if stProduitpgi = 'GRF' then
   begin
   SetControlProperty ('RAG_TYPEACTION','Datatype','RTTYPEACTIONFOU');
   SetControlVisible ('TRAC_TABLELIBRE1',False); SetControlVisible ('RAG_TABLELIBRE1',False);
   SetControlVisible ('TRAC_TABLELIBRE2',False); SetControlVisible ('RAG_TABLELIBRE2',False);
   SetControlVisible ('TRAC_TABLELIBRE3',False); SetControlVisible ('RAG_TABLELIBRE3',False);
   soRtactgestech:=GetParamsocSecur('SO_RFACTGESTECH',False);
   soRtactgestcout:=GetParamsocSecur('SO_RFACTGESTCOUT',False);
   end
else
   begin
   SetControlVisible ('TRAC_TABLELIBREF1',False); SetControlVisible ('RAG_TABLELIBREF1',False);
   SetControlVisible ('TRAC_TABLELIBREF2',False); SetControlVisible ('RAG_TABLELIBREF2',False);
   SetControlVisible ('TRAC_TABLELIBREF3',False); SetControlVisible ('RAG_TABLELIBREF3',False);
   soRtactgestech:=GetParamsocSecur('SO_RTACTGESTECH',False);
   soRtactgestcout:=GetParamsocSecur('SO_RTACTGESTCOUT',False);
   end;
if soRtactgestech = FALSE then
   begin
   SetControlEnabled('RAG_DATEECHEANCE',FALSE) ;
   SetControlEnabled('TRAG_DATEECHEANCE',FALSE) ;
   end;
if soRtactgestcout = FALSE then
   begin
   SetControlEnabled('RAG_COUTACTION',FALSE) ;
   SetControlEnabled('TRAG_COUTACTION',FALSE) ;
   end;
SetControlText ('PRODUITPGI',stProduitpgi);
THEdit(GetControl('RAG_DATEACTION')).OnExit := OnExit_DateActionGen;
end;

procedure TOM_ACTIONSGENERIQUES.OnDeleteRecord  ;
begin
Inherited ;

if ExisteSQL('SELECT RAC_LIBELLE FROM ACTIONS WHERE RAC_NUMACTGEN='
     + IntToStr(GetField('RAG_NUMACTGEN'))+' AND RAC_OPERATION="'+codeoperation+'"') then
   BEGIN
   LastError:=3;
   LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
   exit ;
   end ;
end;

procedure TOM_ACTIONSGENERIQUES.ChgtTypeActionGen;
var Affect : Boolean;
begin
if (GetField ('RAG_TYPEACTION') <> '') and (GetField ('RAG_TYPEACTION') <> Old_Typeaction) then
   begin
   Old_Typeaction := GetField ('RAG_TYPEACTION');
   Affect := True;
   LookEcranActGen(Affect);
   end;
end;

procedure TOM_ACTIONSGENERIQUES.LookEcranActGen (Affect : Boolean);
var TobTypActEncours : Tob;
begin
   VH_RT.TobTypesAction.Load;

   TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RAG_TYPEACTION'),'---',0],TRUE) ;
   if TobTypActEncours = Nil then exit;

   if (TobTypActEncours.GetValue('RPA_GESTDATECH') = 'X') then
      begin
      SetControlEnabled('RAG_DATEECHEANCE',TRUE) ;
      SetControlEnabled('TRAG_DATEECHEANCE',TRUE) ;
      if (Affect) and (TobTypActEncours.GetValue('RPA_DELAIDATECH') <> 0) then
          SetField('RAG_DATEECHEANCE', RTCalculEch(GetField('RAG_DATEACTION'),StrToInt(TobTypActEncours.GetValue('RPA_DELAIDATECH')),TobTypActEncours.GetValue('RPA_WEEKEND')));
      end
   else
      begin
      SetControlEnabled('RAG_DATEECHEANCE',FALSE) ;
      SetControlEnabled('TRAG_DATEECHEANCE',FALSE) ;
      if GetField('RAG_DATEECHEANCE') <> iDate2099 then  SetField('RAG_DATEECHEANCE', iDate2099);
      end;
   if (TobTypActEncours.GetValue('RPA_GESTCOUT') = 'X') then
      begin
      SetControlEnabled('RAG_COUTACTION',TRUE) ;
      SetControlEnabled('TRAG_COUTACTION',TRUE) ;
      end
   else
      begin
      SetControlEnabled('RAG_COUTACTION',FALSE) ;
      SetControlEnabled('TRAG_COUTACTION',FALSE) ;
      if GetField('RAG_COUTACTION') <> 0 then  SetField('RAG_COUTACTION', 0);
      end;
   if Affect then
      begin
      SetField ('RAG_TABLELIBRE1',TobTypActEncours.GetValue('RPA_TABLELIBRE1'));
      SetField ('RAG_TABLELIBRE2',TobTypActEncours.GetValue('RPA_TABLELIBRE2'));
      SetField ('RAG_TABLELIBRE3',TobTypActEncours.GetValue('RPA_TABLELIBRE3'));
      SetField ('RAG_TABLELIBREF1',TobTypActEncours.GetValue('RPA_TABLELIBREF1'));
      SetField ('RAG_TABLELIBREF2',TobTypActEncours.GetValue('RPA_TABLELIBREF2'));
      SetField ('RAG_TABLELIBREF3',TobTypActEncours.GetValue('RPA_TABLELIBREF3'));
      if TobTypActEncours.GetValue('RPA_ETATACTION') <> '' then
        SetField ('RAG_ETATACTION',TobTypActEncours.GetValue('RPA_ETATACTION'))
      else
        SetField('RAG_ETATACTION', Encours);
      end;
   if (TobTypActEncours.GetValue('RPA_REPLICLIB') = 'X' ) then
      SetField ('RAG_LIBELLE',THValComboBox(GetControl('RAG_TYPEACTION')).Text);
end;

procedure TOM_ACTIONSGENERIQUES.OnExit_DateActionGen(Sender: TObject);
var TobTypActEncours : TOB;
begin
if (DS<>Nil) and (DS.State in [dsInsert,dsEdit]) then
  if GetField ('RAG_DATEACTION') <> Old_DateAction then
     begin
     Old_DateAction := GetField ('RAG_DATEACTION');

     VH_RT.TobTypesAction.Load;

     TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RAG_TYPEACTION'),'---',0],TRUE) ;
     if TobTypActEncours <> Nil then
       begin
         if TobTypActEncours.GetValue('RPA_DELAIDATECH') <> 0 then SetField('RAG_DATEECHEANCE',RTCalculEch(GetField('RAG_DATEACTION'),StrToInt(TobTypActEncours.GetValue('RPA_DELAIDATECH')),TobTypActEncours.GetValue('RPA_WEEKEND')));
       end;
     end;
end;

procedure TOF_RTActionsLot.OnArgument (Arguments : String );
var Critere,ChampMul,ValMul : String;
    x : integer;
begin
inherited ;
  StParent:=''; ErreurLot:=0;
  Repeat
      Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
      if Critere<>'' then
      begin
          x:=pos('=',Critere);
          if x<>0 then
          begin
             ChampMul:=copy(Critere,1,x-1);
             ValMul:=copy(Critere,x+1,length(Critere));
             if ChampMul='TIERSCHAINAGE' then
                StParent := ValMul;
             if ChampMul='NUMCHAINAGE' then
                iNumChainage := StrToInt(ValMul);
             if ChampMul='PRODUITPGI' then
                stProduitpgi := ValMul;
          end
          else TypeTrait := Critere;
      end;
  until  Critere='';
  if StParent <> '' then
  begin
    SetControlVisible('RAG_INTERVENANT',false);
    SetControlVisible('TRAG_INTERVENANT',false);
    SetControlVisible('NOMRESPONSABLE',false);
    SetControlVisible('BFerme',true);
  end;
  if (StParent <> '') or (TypeTrait = 'FIC') then
  begin
    SetControlVisible('RAG_BLOCNOTE',false);
    SetControlVisible('TRAG_BLOCNOTE',false);
  end;
  if stProduitpgi = 'GRF' then
     begin
     SetControlProperty ('RAG_TYPEACTION','Datatype','RTTYPEACTIONFOU');
     SetControlVisible ('TRAC_TABLELIBRE1',False); SetControlVisible ('RAG_TABLELIBRE1',False);
     SetControlVisible ('TRAC_TABLELIBRE2',False); SetControlVisible ('RAG_TABLELIBRE2',False);
     SetControlVisible ('TRAC_TABLELIBRE3',False); SetControlVisible ('RAG_TABLELIBRE3',False);
     end
  else
     begin
     SetControlVisible ('TRAC_TABLELIBREF1',False); SetControlVisible ('RAG_TABLELIBREF1',False);
     SetControlVisible ('TRAC_TABLELIBREF2',False); SetControlVisible ('RAG_TABLELIBREF2',False);
     SetControlVisible ('TRAC_TABLELIBREF3',False); SetControlVisible ('RAG_TABLELIBREF3',False);
     end;
THEdit(GetControl('RAG_DATEACTION')).OnExit := OnExit_DateActionLot;
{$Ifdef GIGI}
 if (GetControl('RAG_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RAG_OPERATION',false);
    SetControlVisible('TRAG_OPERATION',false);
    end;
{$endif}
if TypeTrait = 'CLI' then Ecran.Caption := Ecran.Caption + ' ' + TexteMessage[13]
else if TypeTrait = 'CONT' then Ecran.Caption := Ecran.Caption + ' ' + TexteMessage[14]
else if TypeTrait = 'FIC' then Ecran.Caption := Ecran.Caption + ' ' + TexteMessage[15];
UpdateCaption(Ecran);
{$IFDEF GRCLIGHT}
  if not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False) then
    begin
    SetControlVisible('RAG_OPERATION',false);
    SetControlVisible('TRAG_OPERATION',false);
    end;
{$ENDIF GRCLIGHT}
end;

procedure TOF_RTActionsLot.OnLoad;
begin
inherited ;
SetControlText('RAG_DATEECHEANCE', DateToStr(iDate2099));
SetControlText('RAG_ETATACTION', Encours);
SetControlText('RAG_INTERVENANT',VH_RT.RTResponsable);
SetControlText('NOMRESPONSABLE',VH_RT.RTNomResponsable);
Old_Typeaction := '';
SetFocusControl('RAG_TYPEACTION');
SauvBlocnote := GetControlText('RAG_BLOCNOTE');
end;

procedure TOF_RTActionsLot.OnUpdate;
var TobAction : TOB;
    ListeResp : String;
begin
inherited ;
if GetControlText ('RAG_TYPEACTION') = '' then
   begin
   Lasterror:=4; ErreurLot:=1;
   LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
   SetFocusControl('RAG_TYPEACTION') ;
   exit;
   end;

if GetControlText ('RAG_LIBELLE') = '' then
   begin
   Lasterror:=5; ErreurLot:=1;
   LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
   SetFocusControl('RAG_LIBELLE') ;
   exit;
   end;

if StrToDate(GetControlText ('RAG_DATEACTION')) = iDate1900 then
   begin
   Lasterror:=2;  ErreurLot:=1;
   LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
   SetFocusControl('RAG_DATEACTION') ;
   exit;
   end;

if ((stProduitpgi = 'GRC') and (GetParamsocSecur('SO_RTACTGESTECH',False))) or
   ((stProduitpgi = 'GRF') and (GetParamsocSecur('SO_RFACTGESTECH',False))) then
   begin
   if StrToDate(GetControlText ('RAG_DATEECHEANCE')) < StrToDate(GetControlText ('RAG_DATEACTION')) then
      begin
      Lasterror:=1; ErreurLot:=1;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      SetFocusControl('RAG_DATEECHEANCE') ;
      exit;
      end;
   end;

if GetControlText ('RAG_INTERVENANT') <> '' then
   begin
   if (Not ExisteSQL ('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+GetControlText('RAG_INTERVENANT')+'"')) then
      begin
      SetFocusControl('RAG_INTERVENANT');
      LastError := 7; ErreurLot:=1;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit;
      end;
   end;

if ( (stProduitpgi='GRC') and ( not RTDroitModifActions('',GetControlText('RAG_TYPEACTION'),GetControlText('RAG_INTERVENANT')) ) ) or
   ( (stProduitpgi='GRF') and ( not RTDroitModifActionsF('',GetControlText('RAG_TYPEACTION'),GetControlText('RAG_INTERVENANT')) ) ) then
    begin
    SetFocusControl('RAG_TYPEACTION');
    LastError :=8; ErreurLot:=1;
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
    exit;
    end;

if GetControlText ('RAG_OPERATION') <> '' then
   begin
   if (Not ExisteSQL ('SELECT ROP_LIBELLE FROM OPERATIONS WHERE ROP_FERME<>"X" AND ROP_OPERATION="'+GetControlText('RAG_OPERATION')+'" AND ROP_PRODUITPGI="'+stProduitpgi+'"')) then
      begin
      SetFocusControl('RAG_OPERATION');
      LastError := 9; ErreurLot:=1;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit;
      end;
   end;

TobAction:=TOB.create ('ACTIONSGENERIQUES',NIL,-1);
TobAction.PutValue ('RAG_LIBELLE',GetControlText('RAG_LIBELLE'));
TobAction.PutValue ('RAG_TYPEACTION',GetControlText('RAG_TYPEACTION'));
if StParent = '' then
begin
   TobAction.PutValue ('RAG_INTERVENANT',GetControlText('RAG_INTERVENANT'));
   if GetControlText('RAG_BLOCNOTE') <> SauvBlocNote then TobAction.PutValue ('RAG_BLOCNOTE',GetControlText('RAG_BLOCNOTE'));
end
else
begin
   TobAction.PutValue ('RAG_INTERVENANT',StParent);
   if LaTob <> Nil then
      begin
        TobAction.PutValue ('RAG_BLOCNOTE',LaTob.GetValue('RAC_BLOCNOTE'));
      TobAction.PutValue ('RAG_AFFAIRE',LaTob.GetValue('RAC_AFFAIRE'));
      TobAction.PutValue ('RAG_AFFAIRE0',LaTob.GetValue('RAC_AFFAIRE0'));
      TobAction.PutValue ('RAG_AFFAIRE1',LaTob.GetValue('RAC_AFFAIRE1'));
      TobAction.PutValue ('RAG_AFFAIRE2',LaTob.GetValue('RAC_AFFAIRE2'));
      TobAction.PutValue ('RAG_AFFAIRE3',LaTob.GetValue('RAC_AFFAIRE3'));
      TobAction.PutValue ('RAG_AVENANT',LaTob.GetValue('RAC_AVENANT'));
      end;
end;
if iNumChainage <> 0 then
   TobAction.PutValue ('RAG_NUMACTGEN',iNumChainage);
TobAction.PutValue ('RAG_DATEACTION',StrToDate(GetControlText('RAG_DATEACTION')));
TobAction.PutValue ('RAG_DATEECHEANCE',StrToDate(GetControlText('RAG_DATEECHEANCE')));
TobAction.PutValue ('RAG_ETATACTION',GetControlText('RAG_ETATACTION'));
TobAction.PutValue ('RAG_COUTACTION',GetControlText('RAG_COUTACTION'));
if stProduitpgi = 'GRF' then
  begin
  TobAction.PutValue ('RAG_TABLELIBREF1',GetControlText('RAG_TABLELIBREF1'));
  TobAction.PutValue ('RAG_TABLELIBREF2',GetControlText('RAG_TABLELIBREF2'));
  TobAction.PutValue ('RAG_TABLELIBREF3',GetControlText('RAG_TABLELIBREF3'));
  end
else
  begin
  TobAction.PutValue ('RAG_TABLELIBRE1',GetControlText('RAG_TABLELIBRE1'));
  TobAction.PutValue ('RAG_TABLELIBRE2',GetControlText('RAG_TABLELIBRE2'));
  TobAction.PutValue ('RAG_TABLELIBRE3',GetControlText('RAG_TABLELIBRE3'));
  end;

TobAction.PutValue ('RAG_OPERATION',GetControlText('RAG_OPERATION'));
TobAction.PutValue ('RAG_PRODUITPGI',stProduitpgi);
TobAction.AddChampSupValeur ('RAG_DUREEACTION',0);
if dureeAct <> 0 then
  TobAction.PutValue ('RAG_DUREEACTION',dureeAct);
TheTob := TobAction;
if StParent <> '' then
   ListeResp:=AGLLanceFiche ('RT','RTRESSOURCE_LOT','','','')
else
if TypeTrait = 'CLI' then AGLLanceFiche ('RT','RTPROS_SELECT_LOT','','','PRODUITPGI='+stProduitpgi)
else if TypeTrait = 'CONT' then AGLLanceFiche ('RT','RTCONT_SELECT_LOT','','','PRODUITPGI='+stProduitpgi)
else AGLLanceFiche ('RT','RTPARAMFIC','','','');
TobAction.free;
TFVierge(Ecran).Retour:=ListeResp;
end;

procedure TOF_RTActionsLot.OnClose ;
begin
  if ErreurLot <> 0 then
  begin
    ErreurLot:=0;
    LastError:=1;
    exit;
  end;
end;

procedure TOF_RTActionsLot.ChgtTypeActionLot;
var Affect : Boolean;
begin
if (GetControltext ('RAG_TYPEACTION') <> '') and (GetControltext ('RAG_TYPEACTION') <> Old_Typeaction) then
   begin
   Old_Typeaction := GetControltext ('RAG_TYPEACTION');
   Affect := True;
   LookEcranActLot(Affect);
   end;
end;

procedure TOF_RTActionsLot.LookEcranActLot (Affect : Boolean);
var TobTypActEncours : Tob;
begin
   VH_RT.TobTypesAction.Load;

   TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetControltext('RAG_TYPEACTION'),'---',0],TRUE) ;
   if TobTypActEncours = Nil then exit;

   if (TobTypActEncours.GetValue('RPA_GESTDATECH') = 'X') then
      begin
      SetControlEnabled('RAG_DATEECHEANCE',TRUE) ;
      SetControlEnabled('TRAG_DATEECHEANCE',TRUE) ;
      if (Affect) and (TobTypActEncours.GetValue('RPA_DELAIDATECH') <> 0) then
          SetControlText('RAG_DATEECHEANCE', DateToStr(RTCalculEch(StrToDate(GetControlText('RAG_DATEACTION')),StrToInt(TobTypActEncours.GetValue('RPA_DELAIDATECH')),TobTypActEncours.GetValue('RPA_WEEKEND'))));
      end
   else
      begin
      SetControlEnabled('RAG_DATEECHEANCE',FALSE) ;
      SetControlEnabled('TRAG_DATEECHEANCE',FALSE) ;
      if StrToDate(GetControlText('RAG_DATEECHEANCE')) <> iDate2099 then  SetControlText('RAG_DATEECHEANCE', DateToStr(iDate2099));
      end;
   if (TobTypActEncours.GetValue('RPA_GESTCOUT') = 'X') then
      begin
      SetControlEnabled('RAG_COUTACTION',TRUE) ;
      SetControlEnabled('TRAG_COUTACTION',TRUE) ;
      end
   else
      begin
      SetControlEnabled('RAG_COUTACTION',FALSE) ;
      SetControlEnabled('TRAG_COUTACTION',FALSE) ;
      if Valeur(GetControlText('RAG_COUTACTION')) <> 0 then  SetControlText('RAG_COUTACTION', FloatToStr(0));
      end;
   if (Affect) and (stProduitpgi='GRF') then
      begin
      SetControlText ('RAG_TABLELIBREF1',TobTypActEncours.GetValue('RPA_TABLELIBREF1'));
      SetControlText ('RAG_TABLELIBREF2',TobTypActEncours.GetValue('RPA_TABLELIBREF2'));
      SetControlText ('RAG_TABLELIBREF3',TobTypActEncours.GetValue('RPA_TABLELIBREF3'));
      end;
   if (Affect) and (stProduitpgi='GRC') then
      begin
      SetControlText ('RAG_TABLELIBRE1',TobTypActEncours.GetValue('RPA_TABLELIBRE1'));
      SetControlText ('RAG_TABLELIBRE2',TobTypActEncours.GetValue('RPA_TABLELIBRE2'));
      SetControlText ('RAG_TABLELIBRE3',TobTypActEncours.GetValue('RPA_TABLELIBRE3'));
      end;

   if (TobTypActEncours.GetValue('RPA_REPLICLIB') = 'X' ) then
      SetControlText ('RAG_LIBELLE',THValComboBox(GetControl('RAG_TYPEACTION')).Text);

   if Affect and (not VarIsNull (TobTypActEncours.GetValue('RPA_DUREEACTION')) ) then
     dureeAct:=TobTypActEncours.GetValue('RPA_DUREEACTION');
   if Affect and (not VarIsNull (TobTypActEncours.GetValue('RPA_ETATACTION')) ) then
     SetControlText ('RAG_ETATACTION',TobTypActEncours.GetValue('RPA_ETATACTION'));
end;

procedure TOF_RTActionsLot.OnExit_DateActionLot(Sender: TObject);
var TobTypActEncours : TOB;
begin
  VH_RT.TobTypesAction.Load;

  TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetControltext('RAG_TYPEACTION'),'---',0],TRUE) ;
  if TobTypActEncours <> Nil then
  begin
    if TobTypActEncours.GetValue('RPA_DELAIDATECH') <> 0 then SetControlText('RAG_DATEECHEANCE',DateToStr(RTCalculEch(StrToDate(GetControlText('RAG_DATEACTION')),StrToInt(TobTypActEncours.GetValue('RPA_DELAIDATECH')),TobTypActEncours.GetValue('RPA_WEEKEND'))));
  end;
end;

//  Fonctions de générations des actions génériques
// (à partir de la fiche RTPROS_MUL_SELECT ou RTPRO_SELECT_CONT)
// (à partir de la fiche RTPROS_SELECT_LOT ou RTCONT_SELECT_LOT)
procedure TOF_RtSelectProsp.OnLoad;
var xx_where,StWhere,StListeActions,StJoint,ListeCombos,Confid : string;
    NumAct : THValComboBox;
    DateDeb,DateFin : TDateTime;
    i: integer;
begin
inherited ;
if (TFMul(Ecran).name <> 'RTPROS_SELECT_LOT') and (TFMul(Ecran).name <> 'RTCONT_SELECT_LOT')then
    begin
    NumAct := THValComboBox(GetControl('NUMACTGEN'));
    numAct.Plus := GetControlText ('OPERATION');
    numAct.DataType := 'RTACTGEN';

    StWhere := '';
    if (GetControlText('SANSSELECT') <> 'X') then
       begin
       DateDeb := StrToDate(GetControlText('DATEACTION'));
       DateFin := StrToDate(GetControlText('DATEACTION_'));
       if GetControlText('SANS') = 'X' then StWhere := 'NOT ';
       if (TFMul(Ecran).name = 'RTPRO_SELECT_CONT') then
          begin
 {$ifdef GIGI} //MCD 04/07/2005
          StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = AFDPCONTTIERS.T_AUXILIAIRE';
          StWhere:=StWhere + ' AND RAC_NUMEROCONTACT = AFDPCONTTIERS.C_NUMEROCONTACT';
 {$else}
          StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RTCONTACTSTIERS.T_AUXILIAIRE';
          StWhere:=StWhere + ' AND RAC_NUMEROCONTACT = RTCONTACTSTIERS.C_NUMEROCONTACT';
 {$ENDIF GIGI}
          end
       else
           begin
           if stProduitpgi = 'GRF' then StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RFFOURNISSEURS.T_AUXILIAIRE'
 {$ifdef GIGI}  //MCD 04/07/2005
           else StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = AFDPTIERS.T_AUXILIAIRE';
 {$else}
           else StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RTTIERS.T_AUXILIAIRE';
 {$ENDIF GIGI}
           end;
       if (GetControlText('TYPEACTION') <> '') and (GetControlText('TYPEACTION') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('TYPEACTION'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND RAC_TYPEACTION in '+StListeActions;
          end;
       if GetControlText('RESPONSABLE') <> '' then
          begin
          ListeCombos := GetControlText('RESPONSABLE');
          if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
          ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
          ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
          StWhere:=StWhere + ' AND RAC_INTERVENANT in '+ListeCombos;
          end;
       for i:=1 to 3 do
          begin
          if (GetControlText('TABLELIBRE'+intToStr(i)) <> '') and (GetControlText('TABLELIBRE'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
              begin
              ListeCombos:=FindEtReplace(GetControlText('TABLELIBRE'+intToStr(i)),';','","',True);
              ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
              if stProduitpgi = 'GRF' then StWhere := StWhere + ' AND RAC_TABLELIBREF' +intToStr(i) +' in ' + ListeCombos
              else StWhere := StWhere + ' AND RAC_TABLELIBRE' +intToStr(i) +' in ' + ListeCombos;
              end;
          end;
       if GetControlText('OPERATIONACT') <> '' then
          begin
          StWhere:=StWhere + ' AND RAC_OPERATION = "'+GetControlText('OPERATIONACT')+'"';
          end;
       if (GetControlText('ETATACTION') <> '') and (GetControlText('ETATACTION') <> TraduireMemoire('<<Tous>>')) then
          begin
          ListeCombos:=FindEtReplace(GetControlText('ETATACTION'),';','","',True);
          ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
          StWhere:=StWhere + ' AND RAC_ETATACTION in '+ListeCombos;
          end;
       if (GetControlText('NIVIMP') <> '') and (GetControlText('NIVIMP') <> TraduireMemoire('<<Tous>>')) then
          begin
          ListeCombos:=FindEtReplace(GetControlText('NIVIMP'),';','","',True);
          ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
          StWhere:=StWhere + ' AND RAC_NIVIMP in '+ListeCombos;
          end;    
       StWhere := StWhere + ' AND RAC_DATEACTION >= "'+UsDateTime(DateDeb) +'" AND RAC_DATEACTION <= "'+UsDateTime(DateFin)+'"))';
       if (GetControlText('XX_WHERE') = '') then
          SetControlText('XX_WHERE',StWhere)
       else
           begin
           xx_where := GetControlText('XX_WHERE');
           xx_where := xx_where + ' and (' + StWhere + ')';
           SetControlText('XX_WHERE',xx_where) ;
           end;
       end;
    end;
{ sélection sur critères lignes/articles }
if (TFMul(Ecran).name = 'RTLIGN_MUL_SELECT') or (TFMul(Ecran).name = 'RFLIGN_MUL_SELECT') then
    begin
       StWhere := '';
       DateDeb := StrToDate(GetControlText('DATEPIECE'));
       DateFin := StrToDate(GetControlText('DATEPIECE_'));
       if GetControlText('PASDEPIECE') = 'X' then StWhere := 'NOT ';
       if stProduitpgi = 'GRF' then StWhere:=StWhere + 'EXISTS (SELECT GL_NUMERO FROM LIGNE WHERE (GL_TIERS = RFFOURNISSEURS.T_TIERS'
       else StWhere:=StWhere + 'EXISTS (SELECT GL_NUMERO FROM LIGNE WHERE (GL_TIERS = RTTIERS.T_TIERS';
       // NATUREPIECEG,ETABLISSEMENT,DEPOT,REPRESENTANT,ARTICLE,TYPEARTICLE,DOMAINE,
       if (GetControlText('ETABLISSEMENT') <> '') and (GetControlText('ETABLISSEMENT') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('ETABLISSEMENT'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_ETABLISSEMENT in '+StListeActions;
          end;
       if (GetControlText('NATUREPIECEG') <> '') and (GetControlText('NATUREPIECEG') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('NATUREPIECEG'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_NATUREPIECEG in '+StListeActions;
          end;
       if (GetControlText('DEPOT') <> '') and (GetControlText('DEPOT') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('DEPOT'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_DEPOT in '+StListeActions;
          end;
       if (GetControlText('TYPEARTICLE') <> '') and (GetControlText('TYPEARTICLE') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('TYPEARTICLE'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_TYPEARTICLE in '+StListeActions;
          end;
       if (GetControlText('DOMAINE') <> '') and (GetControlText('DOMAINE') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('DOMAINE'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_DOMAINE in '+StListeActions;
          end;
       // COLLECTION, FAMILLENIV1,FAMILLENIV2,FAMILLENIV3
       if (GetControlText('COLLECTION') <> '') and (GetControlText('COLLECTION') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('COLLECTION'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_COLLECTION in '+StListeActions;
          end;
       if (GetControlText('FAMILLENIV1') <> '') and (GetControlText('FAMILLENIV1') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('FAMILLENIV1'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_FAMILLENIV1 in '+StListeActions;
          end;
       if (GetControlText('FAMILLENIV2') <> '') and (GetControlText('FAMILLENIV2') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('FAMILLENIV2'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_FAMILLENIV2 in '+StListeActions;
          end;
       if (GetControlText('FAMILLENIV3') <> '') and (GetControlText('FAMILLENIV3') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('FAMILLENIV3'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_FAMILLENIV3 in '+StListeActions;
          end;
       if GetControlText('ARTICLE') <> '' then
          begin
          StWhere:=StWhere + ' AND GL_CODEARTICLE = "'+GetControlText('ARTICLE')+'"';
          end;
       if (stProduitpgi <> 'GRF') and (GetControlText('REPRESENTANT') <> '') then
          begin
          StWhere:=StWhere + ' AND GL_REPRESENTANT = "'+GetControlText('REPRESENTANT')+'"';
          end;
       for i:=1 to 9 do
          begin
          if (GetControlText('LIBREART'+intToStr(i)) <> '') and (GetControlText('LIBREART'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
              begin
              ListeCombos := GetControlText('LIBREART'+intToStr(i));
              if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
              ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
              ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
              StWhere:=StWhere + ' AND GL_LIBREART'+intToStr(i)+' in '+ListeCombos;
              end;
          end;
       if (GetControlText('LIBREARTA') <> '') and (GetControlText('LIBREARTA') <> TraduireMemoire('<<Tous>>')) then
          begin
          ListeCombos := GetControlText('LIBREARTA');
          if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
          ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
          ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
          StWhere:=StWhere + ' AND GL_LIBREARTA in '+ListeCombos;
          end;
       StWhere := StWhere + ' AND GL_DATEPIECE >= "'+UsDateTime(DateDeb) +'" AND GL_DATEPIECE <= "'+UsDateTime(DateFin)+'"))';
       if (GetControlText('XX_WHERE') = '') then
          SetControlText('XX_WHERE',StWhere)
       else
           begin
           xx_where := GetControlText('XX_WHERE');
           xx_where := xx_where + ' and (' + StWhere + ')';
           SetControlText('XX_WHERE',xx_where) ;
           end;
    end;

  if (stProduitpgi <> 'GRF') then Confid:='CON' else Confid:='CONF';
  if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',RTXXWhereConfident(Confid,true))
  else
      begin
      xx_where := GetControlText('XX_WHERE');
      xx_where := xx_where + RTXXWhereConfident(Confid,true);
      SetControlText('XX_WHERE',xx_where) ;
      end;

if (TFMul(Ecran).name = 'RTPROS_MUL_SELECT') or (TFMul(Ecran).name = 'RTLIGN_MUL_SELECT') or
   (TFMul(Ecran).name = 'RFPROS_MUL_SELECT') or (TFMul(Ecran).name = 'RFLIGN_MUL_SELECT') then
    begin
    StJoint := '';
    StJoint := StJoint + 'C_FERME <> "X" and ';
    if (GetCheckBoxState ('PRINCIPAL') <> cbGrayed) then
       StJoint := StJoint + 'C_PRINCIPAL = "' + GetControlText ('PRINCIPAL') + '" and ';
    if (GetCheckBoxState ('PUBLIPOSTAGE') <> cbGrayed) then
       StJoint := StJoint + 'C_PUBLIPOSTAGE = "' + GetControlText ('PUBLIPOSTAGE') + '" and ';
    if (GetControlText('FONCTIONCODEE') <> '') and (GetControlText('FONCTIONCODEE') <> TraduireMemoire('<<Tous>>')) then
       begin
//       ListeCombos:=FindEtReplace(GetControlText('FONCTIONCODEE'),';','","',True);
//       ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
//       StJoint := StJoint + 'C_FONCTIONCODEE in ' + ListeCombos + ' and ';
         StJoint := StJoint + 'C_FONCTIONCODEE = "' + GetControlText ('FONCTIONCODEE') + '" and ';
       end;
    if (GetControlText('SERVICECODE') <> '') and (GetControlText('SERVICECODE') <> TraduireMemoire('<<Tous>>')) then
       begin
         StJoint := StJoint + 'C_SERVICECODE = "' + GetControlText ('SERVICECODE') + '" and ';
       end;
    For i:=1 to 10 do
        begin
        if i < 10 then
           begin
           if (GetControlText('LIBRECONTACT'+intToStr(i)) <> '') and (GetControlText('LIBRECONTACT'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
              begin
    //            ListeCombos:=FindEtReplace(GetControlText('LIBRECONTACT'+intToStr(i)),';','","',True);
    //            ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
    //            StJoint := StJoint + 'C_LIBRECONTACT' +intToStr(i) +' in ' + ListeCombos + ' and ';
              StJoint := StJoint + 'C_LIBRECONTACT' +intToStr(i) +' = "' + GetControlText('LIBRECONTACT'+intToStr(i)) + '" and ';
              end;
           end else
           begin
           if (GetControlText('LIBRECONTACTA') <> '') and (GetControlText('LIBRECONTACTA') <> TraduireMemoire('<<Tous>>')) then
              begin
              StJoint := StJoint + 'C_LIBRECONTACTA' +' = "' + GetControlText('LIBRECONTACTA') + '" and ';
              end;
           end;
        end;
    SetControlText ('XX_JOIN',StJoint);
    end;
end;

{$ifdef GIGI}
//mcd 29/11/2005 pour passer outre ce qui est fait dans latof,
// car pas même vue en GI
procedure TOF_RtSelectProsp.BChercheClick(Sender: TObject);
var stWhere : string;
begin
if not Assigned(GetControl('EXISTE')) then
  begin
  TFMul(Ecran).BCherCheClick(self) ;
	exit;	// mcd 14/02/2006 le champ n'existe pas dans toutes les fiches auant cette tof
	end;
if TCheckBox(GetControl('EXISTE')).state=cbgrayed then
   stwhere:=''
 else
   begin
   if TCheckBox(GetControl('EXISTE')).checked=False then stwhere:='not '
   else  stwhere:='';

   if TfMul(Ecran).name='RTPRO_SELECT_CONT' then     //mcd 09/06/06 si sur contact, la vue n'est pas la même
      stwhere:=stwhere+'exists (select RAC_NUMACTGEN from ACTIONS  where ((RAC_AUXILIAIRE=AFDPCONTTIERS.T_AUXILIAIRE) AND RAC_OPERATION="'
        +GetControlText ('OPERATION')+'" AND (RAC_NUMACTGEN='
        +GetControlText('NUMACTGEN')+')))'
   else stwhere:=stwhere+'exists (select RAC_NUMACTGEN from ACTIONS  where ((RAC_AUXILIAIRE=AFDPTIERS.T_AUXILIAIRE) AND RAC_OPERATION="'
        +GetControlText ('OPERATION')+'" AND (RAC_NUMACTGEN='
        +GetControlText('NUMACTGEN')+')))';
   end;
SetControlText ('XX_WHERE',stwhere);
TFMul(Ecran).BCherCheClick(self) ;
end;
{$endif}

procedure TOF_RtSelectProsp.OnArgument (Arguments : String );
var Critere,ChampMul,ValMul : string;
    x : integer;
    Texiste : TCheckBox;
    F : TForm;
begin
	fMulDeTraitement := true;

inherited ;
StOperation := '';
stProduitpgi := '';
noactgen := 0;
MajContact := True;
if (TFMul(Ecran).name = 'RTPROS_SELECT_LOT') {or (TFMul(Ecran).name = 'RTLIGN_MUL_SELECT')} then
      MajContact := False;
{$ifdef GIGI}   //liste modifier pour la GI, et n'a pas le contact mcd 08/09/2005
   if (TFMul(Ecran).name = 'RTPROS_MUL_SELECT') then MajContact := False;
{$endif GIGI}
  Repeat
      Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
      if Critere <> '' then
      begin
          x := pos('=',Critere);
          if x <> 0 then
          begin
             ChampMul:=copy(Critere,1,x-1);
             ValMul:=copy(Critere,x+1,length(Critere));
             if ChampMul='CODEOPER' then
             begin
                Stoperation := ValMul;
             end;
             if ChampMul='ACTGEN' then
             begin
                noactgen := strToInt(ValMul);
             end;
             if ChampMul='WHERE' then
             begin
             Texiste := TCheckBox(GetControl('EXISTE'));
             Texiste.state :=cbChecked;
             SetControlText('EXISTE','X') ;
             SetControlEnabled('BOuvrir',False) ;
             SetControlEnabled('bSelectAll',False) ;
             SetControlText('XX_WHERE',ValMul);
             SetControlEnabled('NUMACTGEN',False) ;
             end;
             if ChampMul='PRODUITPGI' then
                stProduitpgi := ValMul;
          end;
      end;
  until  Critere='';
F := TForm (Ecran);
if stProduitpgi = 'GRF' then
  begin
  if GetParamSocSecur('SO_RTGESTINFOS003',False) = True then
      MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
  end
else MulCreerPagesCL(F,'NOMFIC=GCTIERS');

// Paramétrage des libellés des familles, stat. article et dimensions
if (TFMul(Ecran).name = 'RTLIGN_MUL_SELECT') or (TFMul(Ecran).name = 'RFLIGN_MUL_SELECT') then
begin
  for x:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+InttoStr(x),Ecran);
  ChangeLibre2('TGA_COLLECTION',Ecran);
  MajChampsLibresArticle(TForm(Ecran));
  FiltreComboTypeArticle(THMultiValComboBox(GetControl('TYPEARTICLE')));
end;
{ Onglets infos complémentaires des contacts sur les 3 traitements
  pour contact et contact des affaires }
if (GetParamSocSecur('SO_RTGESTINFOS006',False) = True) and ( (TFMul(Ecran).name = 'RTPRO_SELECT_CONT')
    or (TFMul(Ecran).name = 'RTCONT_SELECT_LOT') ) then
    MulCreerPagesCL(F,'NOMFIC=YYCONTACT');
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    if not (ctxscot in V_PGI.PGICOntexte) then
       begin
       SetControlVisible ('T_MOISCLOTURE',false);
       SetControlVisible ('T_MOISCLOTURE_',false);
       SetControlVisible ('TT_MOISCLOTURE',false);
       SetControlVisible ('TT_MOISCLOTURE_',false);
       end;
    end;
  end;
{$Ifdef GIGI}
 if (GetControl('T_REPRESENTANT') <> nil) then  SetControlVisible('T_REPRESENTANT',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then  SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI_', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI_', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI_', 'Plus', VH_GC.AfNatTiersGRCGI);
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
      //MCD 04/07/2005
 if ((Ecran).name = 'RTPROS_SELECT_LOT') or ((Ecran).name = 'RTPROS_MUL_SELECT')then
  begin //on peut avoir des info du DP
  // mcd 20/09/2005 agl 148 TfMul(Ecran).dbliste := 'AFMULTIERSLOT';
  TfMul(Ecran).Setdbliste ('AFMULTIERSLOT');
  //mcd 20/09/2005 agl 148 if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste  := 'AFMULTIERSLOT';
  MulCreerPagesDP((ecran));
  end;
 if ((Ecran).name = 'RTCONT_SELECT_LOT') or ((Ecran).name = 'RTPRO_SELECT_CONT') then
  begin //on peut avoir des info du DP
  //mcd 20/09/2005 agl 148 TfMul(Ecran).dbliste := 'AFMULCONTACTIONS';
  TfMul(Ecran).Setdbliste ( 'AFMULCONTACTIONS');
  //mcd 20/09/2005 agl 148 if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste  := 'AFMULCONTACTIONS';
  MulCreerPagesDP(Ecran);
  end;
  TToolBarButton97(GetControl('BCHERCHE')).OnClick := bChercheClick; //mcd 29/11/2005
{$endif}
end;

procedure TOF_RtGenereAct_Fic.OnArgument (Arguments : String );
var stFic : string;
begin
inherited ;
  TobActGen := LaTob;
  ErreurFic := False;
  SetControlText ('NOMFIC',GetSynRegKey('FichierGenerActions',stFic,TRUE));
end;

procedure TOF_RtGenereAct_Fic.OnUpdate;
var stFic,Col,Liste : string;
    TobAction : TOB;
begin
inherited ;
stFic := GetControlText ('NOMFIC');
Col := GetControlText ('COLTIERS');
Liste := GetControlText ('LISTECOL');
if stFic = '' then
   begin
   PGIBox('Le nom de fichier doit être renseigné', ecran.caption);
   SetFocusControl('NOMFIC') ;
   ErreurFic := True;
   exit;
   end;
if not FileExists(stFic) then
   begin
   PGIBox('Le fichier n''existe pas', ecran.caption);
   SetFocusControl('NOMFIC') ;
   ErreurFic := True;
   exit;
   end;
if Valeuri(Col) <= 0 then
   begin
   PGIBox('Le numéro de colonne du tiers doit être renseigné', ecran.caption);
   SetFocusControl('COLTIERS') ;
   ErreurFic := True;
   exit;
   end;
AffColTiers (True);
if ErreurFic = True then exit;
AffColBlocNote (True);
if ErreurFic = True then exit;

if TobActGen = Nil then
  begin
  PGIInfo('Informations de l'' action non trouvées','');
  exit;
  end;
TobAction:=TOB.create ('ACTIONS',NIL,-1);
AlimTobAction (TobAction,TobActGen,'',0);
TraitementFic (Tobaction,TobActGen,stFic,Valeuri(Col),Liste);
TobAction.free;
SaveSynRegKey('FichierGenerActions',stFic,TRUE);
end;

procedure TOF_RtGenereAct_Fic.OnClose ;
begin
inherited ;
if ErreurFic then
   begin
   LastError := 1 ;
   ErreurFic := False;
   end
end;

Procedure TOF_RtGenereAct_Fic.AffColTiers (AffAno : Boolean) ;
var FLig : textfile;
    StLig,Critere,stFic,Col,WStLig : string;
    i,Lg,LgLue,NbCol : integer;
begin
i := 0;
LgLue := 0;
NbCol := 0;
stFic := GetControlText ('NOMFIC');
Col := GetControlText ('COLTIERS');
if not FileExists(stFic) then exit;
AssignFile(FLig, stFic);
Reset (FLig);
readln (FLig,StLig);
CloseFile(FLig);
Lg := length (StLig);
if (Lg = 0) then exit;
WStLig := StLig;
Repeat
Critere:=trim(ReadTokenSt(StLig));
LgLue := LgLue + length (Critere) + 1;
Inc (NbCol);
until  LgLue >= Lg;
if (Valeuri(Col)  <= 0) or (Valeuri(Col) > NbCol) or (Valeuri(Col) > MaxCol) then
   begin
   if AffAno then
     begin
     PGIBox ('Colonne du tiers incorrecte');
     SetFocusControl('COLTIERS') ;
     ErreurFic := True;
     end;
   exit;
   end;
StLig := WStLig;
Repeat
Critere:=trim(ReadTokenSt(StLig));
Inc (i);
until  i = Valeuri(Col);
SetControlText ('CONTENUCOLTIERS',Critere);
end;

Procedure TOF_RtGenereAct_Fic.AffColBlocNote (AffAno : Boolean) ;
var FLig : textfile;
    StLig,Critere,stFic,ColBlocNote,WStLig,WColBlocNote,Liste,Col : string;
    i,Lg,LgLue,NbCol : integer;
    ChampsLignes: array[1..50] of string;
begin
LgLue := 0;
NbCol := 0;
stFic := GetControlText ('NOMFIC');
ColBlocNote := GetControlText ('LISTECOL');
if ColBlocNote = '' then exit;

if not FileExists(stFic) then exit;
AssignFile(FLig, stFic);
Reset (FLig);
readln (FLig,StLig);
CloseFile(FLig);
Lg := length (StLig);
if (Lg = 0) then exit;
WStLig := StLig;
Repeat
Critere:=trim(ReadTokenSt(StLig));
LgLue := LgLue + length (Critere) + 1;
Inc (NbCol);
until  LgLue >= Lg;

WColBlocNote := ColBlocNote;
Repeat
Critere:=trim(ReadTokenSt(ColBlocNote));
if Critere<>'' then
   begin
   if (Valeuri(Critere)  <= 0) or (Valeuri(Critere) > NbCol) or (Valeuri(Critere) > MaxCol) then
     begin
     if AffAno then
        begin
        PGIBox ('Colonnes incorrectes');
        SetFocusControl('LISTECOL') ;
        ErreurFic := True;
        end;
     exit;
     end;
   end;
until  Critere = '';

for i := 1 to MaxCol do
    begin
    ChampsLignes[i] := '';
    end;
StLig := WStLig;
if Nbcol > MaxCol then NbCol := MaxCol;
for i := 1 to NbCol do
  begin
  Critere:=trim(ReadTokenSt(StLig));
  ChampsLignes[i] := Critere;
  end;
ColBlocNote := WColBlocNote;
Liste := '';
Repeat
Col:=trim(ReadTokenSt(ColBlocNote));
if Col<>'' then
   begin
   Liste := Liste + ChampsLignes[Valeuri(Col)] + ' ';
   end;
until  Col='';

SetControlText ('CONTENUCOLBLOCNOTE',Liste);
end;

procedure TOF_RtSelectProsp.GenerationDesActions(POperation : string;PActgen:integer);
var  F : TFMul ;
     codeprospect,codetiers,commercial : string;
     i,nocontact : integer;
     trouve : integer;
{$IFDEF EAGLCLIENT}
       L : THGrid;
{$ELSE}
       L : THDBGrid;
{$ENDIF}
     Q : THQuery;
     QAG : TQuery;
     TobActGen : TOB;
     TobAction : TOB;
begin
F:=TFMul(Ecran);

if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
   begin
   PGIInfo('Aucun élément sélectionné','');
   exit;
   end;
if PGIAsk('Confirmez-vous le traitement ?','')<>mrYes then
   begin
   if F.FListe.AllSelected then
      begin
      F.bSelectAll.Down := False;
      F.FListe.AllSelected:=False;
      end;
   exit ;
   end;

{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
   if not F.Fetchlestous then
     begin
     F.bSelectAllClick(Nil);
     F.bSelectAll.Down := False;
     exit;
     end else
     F.Fliste.AllSelected := true;
{$ENDIF}

L:= F.FListe;
Q:= F.Q;

{$IFDEF AFFAIRE}
if (POperation <> '') then StOperation := POperation;
if (PActgen <> 0) then noactgen := PActgen;
{$ENDIF}
if StOperation = '' then
   begin
   TobActGen := LaTob;
   if TobActGen = Nil then
      begin
       PGIInfo('Informations de l'' action non trouvées','');
       exit;
      end;
   end
else
    begin
    QAG := OpenSQL('Select * from ACTIONSGENERIQUES Where RAG_OPERATION="' +
                 Stoperation + '" AND RAG_NUMACTGEN ='+ IntToStr(noactgen),True,-1, '', True) ;
    trouve:= 0;
    TobActGen := Nil;
    if Not QAG.EOF then
        BEGIN
        TobActGen:=TOB.create ('ACTIONSGENERIQUES',NIL,-1);
        TobActGen.SelectDB('',QAG);
    //TheTob:=TobActGen ;
    //AGLLanceFiche('GC','GCVOIRTOB','','','MONO');
        trouve:=1;
        END ;
    Ferme(QAG) ;
    if (trouve=0) then
       begin
       PGIInfo('Action générique non trouvée','');
       exit;
       end;
    { en attendant la V9 d'avoir le champ dans actions génériques, on le met ds la tob }
    if not TobActGen.FieldExists ('RAG_DUREEACTION') then
      doAlimDureeAction(TobActGen);
    end;

TobAction:=TOB.create ('ACTIONS',NIL,-1);
AlimTobAction (TobAction,TobActGen,Stoperation,noactgen);

if L.AllSelected then
   begin
   InitMove(Q.RecordCount,'');
   Q.First;
   while Not Q.EOF do
      begin
      MoveCur(False);
      codeprospect:=TFmul(Ecran).Q.FindField('T_AUXILIAIRE').asstring ;
      codetiers:=TFmul(Ecran).Q.FindField('T_TIERS').asstring ;
      commercial:=TFmul(Ecran).Q.FindField('T_REPRESENTANT').asstring ;
      nocontact := 0;
      TobAction.PutValue ('RAC_INTERVENANT',TobActGen.GetValue('RAG_INTERVENANT'));
{      if TobActGen.GetValue('RAG_INTERVENANT')<>'' then
         TobAction.PutValue ('RAC_INTERVINT',TobActGen.GetValue('RAG_INTERVENANT')+';')
      else
         TobAction.PutValue ('RAC_INTERVINT','');  }

//      if StOperation <> '' then nocontact := TFmul(Ecran).Q.FindField('C_NUMEROCONTACT').asinteger;
      if MajContact then
      begin
				if TFmul(Ecran).Q.FindField('C_NUMEROCONTACT') = nil then
        begin
      		nocontact := 0;
        end else
        begin
      		nocontact := TFmul(Ecran).Q.FindField('C_NUMEROCONTACT').asinteger;
        end;
      end;
      CreationAction (codeprospect,codetiers,commercial,nocontact,TobAction);
      Q.Next;
      end;
   L.AllSelected:=False;
   end else
   begin
   InitMove(L.NbSelected,'');
   for i:=0 to L.NbSelected-1 do
      begin
      MoveCur(False);
      L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(L.Row-1) ;
{$ENDIF}
      codeprospect:=TFmul(Ecran).Q.FindField('T_AUXILIAIRE').asstring ;
      codetiers:=TFmul(Ecran).Q.FindField('T_TIERS').asstring ;
      commercial:=TFmul(Ecran).Q.FindField('T_REPRESENTANT').asstring ;
      nocontact := 0;
      TobAction.PutValue ('RAC_INTERVENANT',TobActGen.GetValue('RAG_INTERVENANT'));
{      if TobActGen.GetValue('RAG_INTERVENANT')<>'' then
         TobAction.PutValue ('RAC_INTERVINT',TobActGen.GetValue('RAG_INTERVENANT')+';')
      else
         TobAction.PutValue ('RAC_INTERVINT','');  }

//      if StOperation <> '' then nocontact := TFmul(Ecran).Q.FindField('C_NUMEROCONTACT').asinteger;
      if MajContact then
      begin
				if TFmul(Ecran).Q.FindField('C_NUMEROCONTACT') = nil then
        begin
      		nocontact := 0;
        end else
        begin
      		nocontact := TFmul(Ecran).Q.FindField('C_NUMEROCONTACT').asinteger;
        end;
      end;
      CreationAction (codeprospect,codetiers,commercial,nocontact,TobAction);
      end;
   L.ClearSelected;
   end;

if F.bSelectAll.Down then
    F.bSelectAll.Down := False;

FiniMove;
if StOperation <> '' then TobActGen.free;
TobAction.free;

end;

procedure TOF_RtSelectProsp.DoAlimDureeAction(TobActGen : tob);
var TobTypActEncours : Tob;
begin
   TobActGen.AddChampSupValeur('RAG_DUREEACTION',0);
   VH_RT.TobTypesAction.Load;
   TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[TobActGen.GetString('RAG_TYPEACTION'),'---',0],TRUE) ;
   if TobTypActEncours = Nil then exit;
   TobActGen.PutValue('RAG_DUREEACTION',TobTypActEncours.GetValue('RPA_DUREEACTION'));
end;

Function RTLanceFiche_RtActionsLot(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

Procedure AGLAffColTiers( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_RtGenereAct_Fic(totof).AffColTiers(Parms[1]);
end;

Procedure AGLAffColBlocNote( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then ToTof:=TFVierge(F).LaTof else exit;
  TOF_RtGenereAct_Fic(totof).AffColBlocNote(Parms[1]);
end;

Initialization
registerclasses([TOM_ACTIONSGENERIQUES]);
registerclasses([TOF_RtSelectProsp]);
registerclasses([TOF_RtActionsLot]);
registerclasses([TOF_RtGenereAct_Fic]);
RegisterAglProc('GenereAction',FALSE,2,AGLGenereAction);
RegisterAglProc('GenerationDesActions',TRUE,3,AGLGenerationDesActions);
RegisterAglProc('ChgtTypeActionGen', TRUE , 0, AGLChgtTypeActionGen);
RegisterAglProc('ChgtTypeActionLot', TRUE , 0, AGLChgtTypeActionLot);
RegisterAglproc( 'AffColTiers', TRUE ,1 , AGLAffColTiers);
RegisterAglproc( 'AffColBlocNote', TRUE ,1 , AGLAffColBlocNote);
end.
