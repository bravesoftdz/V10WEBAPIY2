{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/11/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTDUPLIMPORT ()
Mots clefs ... : TOF;RTDUPLIMPORT
*****************************************************************}
Unit UtofRTDUPLIMPORT ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_main,
{$else}
     eMul,
     MaineAgl,
{$ENDIF}
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     UTOF ;

Type
  TOF_RTDUPLIMPORT = Class (TOF)
    private
    procedure CodeDupliOnExit ( Sender : Tobject );
    Function ValideCodeDupli : boolean;
    procedure BValideOnClick ( Sender : Tobject );
//    procedure GenereCodeDupli;
  

    public
    CodeDescr       : String;
    TobParSuspect   : Tob;
    TobLig          : Tob;
    CodeDupli       : THEdit;
    LibDupli        : THEdit;
    BValide         : TToolbarButton97;
    BFerme          : TToolbarButton97;


    procedure OnArgument (S : String ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

  procedure LanceDuplImport ( Argument : String );

Implementation

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 27/11/2006
Modifié le ... :   /  /    
Description .. : ouverture de la fenêtre
Mots clefs ... : 
*****************************************************************}
procedure LanceDuplImport ( Argument : String );
begin
  AGLLanceFiche('RT', 'RTDUPLIMPORT', '', '', Argument);
end;


procedure TOF_RTDUPLIMPORT.OnArgument (S : String ) ;
var
  Critere     : String;
  Valeur      : String;
  StrSql      : String;
  StrF        : String;
  Q           : Tquery;

begin
  Inherited ;
  Critere := READTOKENST(S);
  While Critere <> '' do
  begin
    Valeur    := Copy (Critere, pos('=', Critere)+1, length(Critere));
    Critere   := UpperCase( Copy (Critere, 1, pos('=', Critere)-1));
    if Critere = 'CODE' then
      CodeDescr   := Valeur;
    Critere   := READTOKENST(S);
  end;

  CodeDupli   := THEdit(GetControl('CODEDUPLI'));
  LibDupli    := THEdit(GetControl('LIBDUPLI'));
  BValide     := TToolbarButton97(GetControl('BVALIDER'));
  BFerme      := TToolbarButton97(GetControl('BFERME'));

  CodeDupli.OnExit    := CodeDupliOnExit;
  BValide.OnClick     := BValideOnClick;

  TobParsuspect := Tob.Create('PARSUSPECT', nil, -1);
  TobLig := Tob.Create('PARSUSPECTLIG', nil, -1);

  TobParsuspect.SelectDB('"'+CodeDescr+'"', nil);
  StrSql := 'SELECT * FROM PARSUSPECTLIG WHERE RRL_PARSUSPECT="'+CodeDescr+'"';
  Q := OpenSQL(StrSql, True);
  TobLig.LoadDetailDB('PARSUSPECTLIG','','',Q,False);
  Ferme(Q);

  StrF := 'Duplication descriptif : '+CodeDescr;
  TForm(Ecran).Caption := StrF;
  UpdateCaption(TForm(Ecran));

//  GenereCodeDupli;
  LibDupli.Text   := TobParSuspect.GetValue('RSS_LIBELLE');
  CodeDupli.SetFocus;

end ;

procedure TOF_RTDUPLIMPORT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTDUPLIMPORT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RTDUPLIMPORT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RTDUPLIMPORT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_RTDUPLIMPORT.OnClose ;
begin
  Inherited ;
  TobParsuspect.Free;
  TobLig.Free;

end ;

procedure TOF_RTDUPLIMPORT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RTDUPLIMPORT.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_RTDUPLIMPORT.CodeDupliOnExit ( Sender : Tobject );
begin
  ValideCodeDupli;
end;

Function TOF_RTDUPLIMPORT.ValideCodeDupli : Boolean;
var
  StrSql    : String;

begin
  Result := False;
  If CodeDupli.Text = '' then
    CodeDupli.SetFocus;

  StrSql := 'SELECT 1 FROM PARSUSPECT WHERE RSS_PARSUSPECT="'+CodeDupli.text+'"';
  if ExisteSQL(StrSql) then
  begin
    PGIInfo('Ce code de description de fichier existe déjà.', TForm(ecran).Caption);
    Exit;
  end;
  Result := True;
end;


procedure TOF_RTDUPLIMPORT.BValideOnClick(Sender: Tobject);
var
  i         : integer;
  NewCode   : String;

begin
  if not ValideCodeDupli then
  begin
    CodeDupli.SetFocus;
    exit;
  end;

  NewCode := CodeDupli.Text;

  TobParsuspect.PutValue('RSS_PARSUSPECT', NewCode);
  TobParSuspect.PutValue('RSS_LIBELLE', LibDupli.Text);

  for i := 0 to TobLig.Detail.Count-1 do
    TobLig.Detail[i].PutValue('RRL_PARSUSPECT', NewCode);

  begintrans;
  Try
    TobParsuspect.InsertDB(nil);
    for i := 0 to TobLig.Detail.Count-1 do
      TobLig.Detail[i].InsertDB(nil);
    CommitTrans;
  except
    Rollback;
  end;

  TForm(ECran).Close;
end;

{procedure TOF_RTDUPLIMPORT.GenereCodeDupli;
var
  StrCode     : String;
  StrSql      : String;
  Q           : Tquery;

begin
  StrCode   := TobParSuspect.GetValue('RSS_PARSUSPECT');
  if Length(StrCode) <= 8 then
  begin
    StrSql    := 'SELECT MAX(RSS_PARSUSPECT) FROM PARSUSPECT WHERE RSS_PARSUSPECT LIKE "'+StrCode+'_%"';
    Q         := OpenSQL(StrSql, True);

    Ferme(Q);
  end;
end;
}

Initialization
  registerclasses ( [ TOF_RTDUPLIMPORT ] ) ;
end.
