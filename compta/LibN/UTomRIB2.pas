unit UTomRIB;

interface

uses Classes,UTOM,
{$IFDEF EAGLCLIENT}
     utob,eFichList,
{$ELSE}
     db,dbtables,Fiche,FichList,Fe_Main,
{$ENDIF}
     HCtrls,Ent1,HMsgBox,UtilPGI,Controls,HEnt1,SysUtils,
     Hqry ;

Type Tom_RIB = Class (TOM)
  Private
    Auxiliaire : String ;
    FromSaisie : boolean ;
    FirstTime : boolean ;
    StRIB      : String ;
    gtxt_CodeIban1 : THEdit;
    procedure R_PAYSChange(Sender: TObject);
    procedure R_CODEIBANEnter(Sender: TObject);
    procedure R_CODEIBANExit(Sender: TObject);
  Public
    procedure OnNewRecord  ; override ;
    procedure OnUpdateRecord ; override ;
    procedure OnLoadRecord ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument (Arguments : String )  ; override ;
    Function  VerifiRibBic : Boolean ;
    Function  PrincipalUnique : Boolean ;
END ;

const
	// libellés des messages
	TexteMessage: array[1..15] of string 	= (
          {1}        'Vous devez renseigner la domiciliation.',
          {2}        'La domiciliation comporte des caractères interdits.',
          {3}        '3' ,
          {4}        '4' ,
          {5}        '5',
          {6}        'Vous devez renseigner la banque.',
          {7}        'Vous devez renseigner le guichet.',
          {8}        'Vous devez renseigner le numéro de compte.',
          {9}        'Vous devez renseigner la clé du RIB.',
          {10}       'Vous devez renseigner le code BIC.',
          {11}       'Ce numero de Rib ne peut pas être le principal car il n''est pas unique.',
          {12}       '12',
          {13}       '13',
          {14}       '14',
          {15}       '15'
                     );
implementation

procedure Tom_RIB.OnArgument(Arguments: String );
var x : integer;
    critere: string;
    Arg,Val : string;
    cbo_Pays : THValComboBox;
    txt_Iban : THEdit;
begin
inherited ;
Repeat
 Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
 if Critere<>'' then
    begin
    x:=pos('=',Critere);
    if x<>0 then
       begin
       Arg:=copy(Critere,1,x-1);
       Val:=copy(Critere,x+1,length(Critere));
       if Arg='NUMAUX' then Auxiliaire:=Val;
       if Arg='FROMSAISIE' then begin if Val='X' then FromSaisie:=True else FromSaisie:=False; end;
       if Arg='STRIB' then StRIB:=Val;
       //if Arg='ISAUX' then
       end;
    end;
until  Critere='';

if FromSaisie then SetControlChecked('FROMSAISIE',true);
FirstTime:=true;
if ctxPaie in V_PGI.PGIContexte then
   BEGIN
   SetControlVisible('R_SALAIRE',True) ;
//VERSION PAIE S3
{$IFNDEF CCS3}
   SetControlVisible('R_ACOMPTE',True) ;
   SetControlVisible('R_FRAISPROF',True) ;
{$ELSE}
   SetControlChecked('R_SALAIRE', True);
   SetControlEnabled('R_SALAIRE', False);
{$ENDIF}
//FIN VERSION PAIE S3
   END ;
// Code IBAN
cbo_Pays := THValComboBox(GetControl('R_PAYS'));	if (cbo_Pays <> nil) then cbo_Pays.OnChange := R_PAYSChange;
gtxt_CodeIban1 := THEdit(GetControl('CODEIBAN1'));
if (gtxt_CodeIban1 <> nil) then begin
  gtxt_CodeIban1.OnEnter := R_CODEIBANEnter;
  gtxt_CodeIban1.OnExit  := R_CODEIBANExit;
end;
txt_Iban := THEdit(GetControl('CODEIBAN2')); if (txt_Iban <> nil) then txt_Iban.OnExit := R_CODEIBANExit;
txt_Iban := THEdit(GetControl('CODEIBAN3')); if (txt_Iban <> nil) then txt_Iban.OnExit := R_CODEIBANExit;
txt_Iban := THEdit(GetControl('CODEIBAN4')); if (txt_Iban <> nil) then txt_Iban.OnExit := R_CODEIBANExit;
txt_Iban := THEdit(GetControl('CODEIBAN5')); if (txt_Iban <> nil) then txt_Iban.OnExit := R_CODEIBANExit;
txt_Iban := THEdit(GetControl('CODEIBAN6')); if (txt_Iban <> nil) then txt_Iban.OnExit := R_CODEIBANExit;
txt_Iban := THEdit(GetControl('CODEIBAN7')); if (txt_Iban <> nil) then txt_Iban.OnExit := R_CODEIBANExit;
end;

procedure Tom_RIB.OnLoadRecord;
var Etab,Guichet,NumCompte,Cle, Dom : string ;
    TRIB :THTable;
begin
if FromSaisie and FirstTime then
   BEGIN
   FirstTime:=false;
   DecodeRIB(Etab,Guichet,NumCompte,Cle,Dom,StRIB) ; NumCompte:=Trim(NumCompte) ;
   TRIB:=TFFicheliste(Ecran).ta ;
   if TRIB.State=dsBrowse then
      begin
      //TRIB.Locate('R_NUMEROCOMPTE',NumCompte,[]) ;
      while not TRIB.EOF do if uppercase(TRIB.FindField('R_NUMEROCOMPTE').AsString)=uppercase(NumCompte) then Break else TRIB.Next ;
      end ;
   END ;

end;

procedure Tom_RIB.OnUpdateRecord;
Var RibCalcul : String ;
    Reponse : Word ;
    StDom   : String ;
begin
Inherited;
StDom:=GetControlText('R_DOMICILIATION') ;
if StDom='' then
  begin
  SetFocusControl('R_DOMICILIATION'); LastError:=1; LastErrorMsg:=TexteMessage[LastError]; exit;
  end;
if ExisteCarInter(StDom) then
  begin
  SetFocusControl('R_DOMICILIATION'); LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit;
  end;
if not VerifiRibBic then exit;
if not PrincipalUnique then
  begin
  SetFocusControl('R_PRINCIPAL'); LastError:=11; LastErrorMsg:=TexteMessage[LastError]; exit;
  end;
if (VH^.CtrlRIB) and (GetControlText('R_PAYS') = 'FRA') Then
  begin
  RibCalcul:=VerifRib(GetControlText('R_ETABBQ'),GetControlText('R_GUICHET'),GetControlText('R_NUMEROCOMPTE')) ;
  If RibCalcul<>GetControlText('R_CLERIB') Then
     begin
     Reponse:= HShowMessage('1;?caption?;Confirmez-vous la clé du RIB ?;Q;YN;N;N;','','') ;
     if Reponse<>mrYes Then BEGIN SetFocusControl('R_CLERIB'); SetField('R_CLERIB',RibCalcul) ; Exit ; END ;
     end;
  end;
end;

Function Tom_RIB.VerifiRibBic : Boolean ;
BEGIN
Result:=False ;
if GetControlText('R_PAYS')='FRA' then
  BEGIN
   if GetControlText('R_ETABBQ')='' then
      BEGIN SetFocusControl('R_ETABBQ'); LastError:=6; LastErrorMsg:=TexteMessage[LastError]; exit; END ;
   if GetControlText('R_GUICHET')='' then
      BEGIN SetFocusControl('R_GUICHET'); LastError:=7; LastErrorMsg:=TexteMessage[LastError]; exit; END ;
   if GetControlText('R_NUMEROCOMPTE')='' then
      BEGIN SetFocusControl('R_NUMEROCOMPTE'); LastError:=8; LastErrorMsg:=TexteMessage[LastError]; exit; END ;
   if GetControlText('R_CLERIB')='' then
      BEGIN SetFocusControl('R_CLERIB'); LastError:=9; LastErrorMsg:=TexteMessage[LastError]; exit; END ;
  END else
  BEGIN
   if GetControlText('R_CODEBIC')<>''then BEGIN Result:=True ; Exit ; END else
     if ((GetControlText('R_ETABBQ')='')OR(GetControlText('R_GUICHET')='')OR
        (GetControlText('R_NUMEROCOMPTE')='')OR(GetControlText('R_CLERIB')='')) then
        BEGIN SetFocusControl('R_CODEBIC'); LastError:=10; LastErrorMsg:=TexteMessage[LastError]; exit; END ;
  END ;
Result:=True ;
END ;

Function Tom_RIB.PrincipalUnique : Boolean ;
Var Q : TQuery ;
BEGIN
Result:=True ;
if GetControlText('R_PRINCIPAL')<>'X' then Exit ;
Q:=OpenSql('Select R_PRINCIPAL from RIB Where R_AUXILIAIRE="'+Auxiliaire+'" And '+
           'R_PRINCIPAL="X" And R_NUMERORIB<>'+GetControlText('R_NUMERORIB')+'',True) ;
if Not Q.Eof then Result:=False ;
Ferme(Q) ;
END ;

procedure Tom_RIB.OnNewRecord;
Var Q:TQuery;
    LastNum :Longint ;
begin
Inherited;
Q:=OpenSQL('SELECT Max(R_NumeroRib) FROM RIB WHERE R_AUXILIAIRE ="'+Auxiliaire+'"',TRUE) ;
if Not Q.EOF then LastNum:=Q.Fields[0].AsInteger
else LastNum:=0;
SetField('R_AUXILIAIRE',Auxiliaire) ;
SetField('R_NumeroRIB',LastNum+1) ;
SetField('R_PAYS','FRA') ;
SetFocusControl('R_DOMICILIATION') ;
//VERSION PAIE S3
if ctxPaie in V_PGI.PGIContexte then
   BEGIN
{$IFDEF CCS3}
   SetControlChecked('R_SALAIRE', True);
   SetControlEnabled('R_SALAIRE', False);
{$ENDIF}
   END ;
//FIN VERSION PAIE S3
ferme(Q);
end;

procedure Tom_RIB.R_PAYSChange(Sender: TObject);
begin

  // Défaut : France
{  SetControlProperty('TR_ETABBQ','VISIBLE',True);
  SetControlProperty('R_ETABBQ','VISIBLE',True);
  SetControlProperty('TR_GUICHET','VISIBLE',True);
  SetControlProperty('R_GUICHET','VISIBLE',True);
  SetControlProperty('TR_NUMEROCOMPTE','VISIBLE',True);
  SetControlProperty('R_NUMEROCOMPTE','VISIBLE',True);
  SetControlProperty('TR_CLERIB','VISIBLE',True);
  SetControlProperty('R_CLERIB','VISIBLE',True);

  SetControlProperty('TR_ETABBQ','LEFT',21);
  SetControlProperty('R_ETABBQ','LEFT',5);
  SetControlProperty('TR_GUICHET','LEFT',101);
  SetControlProperty('R_GUICHET','LEFT',81);}
  SetControlProperty('TR_NUMEROCOMPTE','LEFT',214);
  SetControlProperty('R_NUMEROCOMPTE','LEFT',165);
  SetControlProperty('TR_CLERIB','LEFT',318);
  SetControlProperty('R_CLERIB','LEFT',307);

  SetControlProperty('R_ETABBQ','MAXLENGTH',5);
  SetControlProperty('R_GUICHET','MAXLENGTH',5);
  SetControlProperty('R_NUMEROCOMPTE','MAXLENGTH',11);
//  SetControlProperty('R_CLERIB','MAXLENGTH',2);

//  SetControlProperty('R_ETABBQ','TABORDER',1);
//  SetControlProperty('R_GUICHET','TABORDER',2);
  SetControlProperty('R_NUMEROCOMPTE','TABORDER',3);
  SetControlProperty('R_CLERIB','TABORDER',4);

  // Espagne
  if (GetControlText('R_PAYS') = 'ESP') then begin
    SetControlProperty('R_ETABBQ','MAXLENGTH',4);
    SetControlProperty('R_GUICHET','MAXLENGTH',4);
    SetControlProperty('R_NUMEROCOMPTE','MAXLENGTH',10);
    SetControlProperty('TR_NUMEROCOMPTE','LEFT',258);
    SetControlProperty('R_NUMEROCOMPTE','LEFT',209);
    SetControlProperty('TR_CLERIB','LEFT',176);
    SetControlProperty('R_CLERIB','LEFT',165);
    SetControlProperty('R_NUMEROCOMPTE','TABORDER',4);
    SetControlProperty('R_CLERIB','TABORDER',3);
{    end
  // Allemagne
  else if (GetControlText('R_PAYS') = 'DEU') then begin
    SetControlProperty('R_ETABBQ','MAXLENGTH',8); // A VOIR
    SetControlProperty('R_NUMEROCOMPTE','MAXLENGTH',10);
    SetControlProperty('TR_ETABBQ','LEFT',221);
    SetControlProperty('R_ETABBQ','LEFT',205);
    SetControlProperty('TR_NUMEROCOMPTE','LEFT',110);
    SetControlProperty('R_NUMEROCOMPTE','LEFT',61);
    SetControlProperty('TR_GUICHET','VISIBLE',False);
    SetControlProperty('R_GUICHET','VISIBLE',False);
    SetControlProperty('TR_CLERIB','VISIBLE',False);
    SetControlProperty('R_CLERIB','VISIBLE',False);
    SetControlProperty('R_NUMEROCOMPTE','TABORDER',1);
}  end;
end;

procedure Tom_RIB.R_CODEIBANEnter(Sender: TObject);
var
	szPays : string;
  i : integer;
  szIban : String;
begin
	szPays := GetField('R_PAYS');
       if szPays = 'FRA' then szPays := 'FR'
  else if szPays = 'ESP' then szPays := 'ES'
  else if szPays = 'DEU' then szPays := 'DE';

    // Pas de pays : vide les THEdit
  if (szPays ='') or (length(szPays) <> 2) then begin
    for i := 1 to 7 do SetControlText('CODEIBAN'+IntToStr(i),'');
    Exit;
  end;

  // Pas de code iban : Le calcul
  if (GetControlText('R_CODEIBAN') = '') then begin
    szIban := calcIBAN(szPays, calcRIB(szPays,GetField('R_ETABBQ'),GetField('R_GUICHET'),GetField('R_NUMEROCOMPTE'),GetField('R_CLERIB')));
    DS.Edit;
    SetControlText('R_CODEIBAN',szIban);
  end;

  // Affiche le code iban par lots de 4
//  if (gtxt_CodeIban1.text = '') then begin
		for i := 1 to 7 do begin
      SetControlText('CODEIBAN'+IntToStr(i),Copy(GetControlText('R_CODEIBAN'),i*4-3,4));
	  end;
//  end;
end;

procedure Tom_RIB.R_CODEIBANExit(Sender: TObject);
var
  i : integer;
  szResult : String;
begin
  // Reconstruit le code iban
  szResult := '';
	for i := 1 to 7 do begin
    szResult := szResult + GetControlText('CODEIBAN'+IntToStr(i));
  end;

  // Renseigne le code iban s'il a changé
  if (GetControlText('R_CODEIBAN') <> szResult) then begin
    DS.Edit;
    SetField('R_CODEIBAN',szResult);
  end;
end;

procedure Tom_RIB.OnChangeField(F: TField);
begin
  inherited;
  if (F.FieldName = 'R_CODEIBAN') then
    R_CODEIBANEnter(nil);
  if (F.FieldName = 'R_PAYS') then
    R_PAYSChange(nil);
end;

Initialization
registerclasses([Tom_RIB]);
end.
