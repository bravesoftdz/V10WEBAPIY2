unit UBTPSequenceCpta;

interface
uses uTob, HEnt1, HCtrls, StdCtrls, Ent1,
     DB,SysUtils,CbpMCD,ADODB,Forms,
    {$IFNDEF DBXPRESS} dbtables {$ELSE} uDbxDataSet {$ENDIF}
;
// COMPTA
Procedure CPInitSequenceCompta;
function CleSequence(TypeSouche,CodeSouche:string;DD:tDateTime;MulTiExo:boolean) : string;
function ExistSequence(Cle: string; Entite : integer=0) : Boolean;
function ReadIncrementSequence(Cle : string ; entity : integer=0) : integer;
function ReadCurrentSequence(Cle: string; entity : integer=0) : integer;
function CreateSequence(Cle: string; valeur : integer; xx : integer=0; entity : integer=0) : Boolean;
function GetNextSequence(Cle: string; nombre : integer; Entity : integer=0) : integer;
// GC - GRC - BTP
procedure GCInitSequenceSouche;
function ConstitueCodeSequenceGC (TypeSouche,CodeSouche : string) : string;


implementation
uses galsystem;

function GetPartage (NomTable : string) : string;
var QQ : TQuery;
begin
  Result := '';
  TRY
    QQ := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="'+NomTable+'"',True,1,'',true);
    if not QQ.eof then
    begin
      Result := QQ.fields[0].AsString;
    end;
    Ferme(QQ);
  EXCEPT
    Exit;
  end;
end;


function wExistTable(Const TableName:String): Boolean;
{$IFDEF V10}
var  Mcd : IMCDServiceCOM;
     Table     : ITableCOM ;
     partage,NomReel : string;
{$ENDIF}
begin
{$IFDEF V10}
  partage := GetPartage(TableName);
  if partage = '' then
  begin
    MCD := TMCD.GetMcd;
    if not mcd.loaded then mcd.WaitLoaded();
    Result := Mcd.TableExists(TableName);
  end else
  begin
    if partage <> '' then NomReel := Partage+'.dbo.'+TableName
                     else NomReel := TableName;
    Result := ISTablePhysExiste (NomReel);
  end;
{$ELSE}
	Result := TableToNum(TableName) <> 0;
{$ENDIF}
end;

function CodeExo(DD:tDateTime):string;
var i : integer;
begin
  Result:=GetEnCours.Code ;
  If (dd>=GetEnCours.Deb) and (dd<=GetEnCours.Fin) then Result:=GetEnCours.Code
  else If (dd>=GetSuivant.Deb) and (dd<=GetSuivant.Fin) then Result:=GetSuivant.Code
  else If (dd>=GetPrecedent.Deb) and (dd<=GetPrecedent.Fin) then Result:=GetPrecedent.Code
  else For i:=1 To 5 do begin
    If (dd>=GetExoClo[i].Deb) And (dd<=GetExoClo[i].Fin)then Result:=GetExoClo[i].Code;
  end;
end;

function Cop(code:string;long:integer):string;
begin
  Result:=Copy(Code+'----',1,long);
end;

function CleSequence(TypeSouche,CodeSouche:string;DD:tDateTime;MulTiExo:boolean) : string;
begin
  if MultiExo then
    Result:=Cop(TypeSouche,3)+Cop(CodeSouche,3)+cop(CodeExo(DD),3)
  else Result:=Cop(TypeSouche,3)+Cop(CodeSouche,3);
end;

function ExistSequence(Cle: string; Entite : integer=0) : Boolean;
var NomPartage,NomReel,NomReel1 : string;
begin
  if wExistTable ('CPSEQCORRESP') then
  begin
    NomPartage := GetPartage('DESEQUENCES');
    if NomPartage <> '' then NomReel := NomPartage+'.dbo.DESEQUENCES'
                        else NomReel := 'DESEQUENCES';
    NomPartage := GetPartage('CPSEQCORRESP');
    if NomPartage <> '' then NomReel1 := NomPartage+'.dbo.CPSEQCORRESP'
                        else NomReel1 := 'CPSEQCORRESP';

	  Result := ExisteSQL('SELECT DSQ_CODE FROM '+NomReel+' WHERE DSQ_CODE=(SELECT CSC_SEQUENCE FROM '+NomReel1+' WHERE CSC_METIER="'+Cle+'")');
  end else
  begin
    NomPartage := GetPartage('DESEQUENCES');
    if NomPartage <> '' then NomReel := NomPartage+'.dbo.DESEQUENCES'
                        else NomReel := 'DESEQUENCES';
	  Result := ExisteSQL('SELECT DSQ_CODE FROM '+NomReel+' WHERE DSQ_CODE="'+Cle+'"');
  end;
end;


function ReadIncrementSequence(Cle : string ; entity : integer=0) : integer;
var QQ : Tquery;
    SQL : string;
begin
  Result := -1;
  try
    if wExistTable ('CPSEQCORRESP') then
    begin
      SQL := 'SELECT DSQ_VALEUR,DSQ_INCREMENT FROM DESEQUENCES WHERE DSQ_CODE=(SELECT CSC_SEQUENCE FROM CPSEQCORRESP WHERE CSC_METIER="'+Cle+'")';
    end else
    begin
      SQL := 'SELECT DSQ_VALEUR,DSQ_INCREMENT FROM DESEQUENCES WHERE DSQ_CODE="'+Cle+'"';
    end;
    //
  	QQ := OpenSql (SQL,True,1,'',true);
    if not QQ.eof then
    begin
      Result := QQ.findField('DSQ_INCREMENT').AsInteger;
    end;
  finally
    ferme (QQ);
  end;
end;

function ReadCurrentSequence(Cle: string; entity : integer=0) : integer;
var QQ : Tquery;
    SQL  : string;
begin
  Result := -1;
  try
    if wExistTable ('CPSEQCORRESP') then
    begin
      SQL := 'SELECT DSQ_VALEUR FROM DESEQUENCES WHERE DSQ_CODE=(SELECT CSC_SEQUENCE FROM CPSEQCORRESP WHERE CSC_METIER="'+Cle+'")';
    end else
    begin
      SQL := 'SELECT DSQ_VALEUR FROM DESEQUENCES WHERE DSQ_CODE="'+Cle+'"';
    end;
  	QQ := OpenSql (SQL,True,1,'',true);
    if not QQ.eof then
    begin
      Result := QQ.findField('DSQ_VALEUR').AsInteger;
    end;
  finally
    ferme (QQ);
  end;
end;

function GetNextSequence(Cle: string; nombre : integer; Entity : integer=0) : integer;
var CNX : TADOConnection;
    QQ : TADOQuery;
    NbRec,OldValue : integer;
    SQL,NomPartage,Nomreel,NomReel1 : string;
begin
  Result := -1;
  CNX := TADOConnection.Create(application);
  Cnx.ConnectionString :=DBSOC.ConnectionString;
  CNX.LoginPrompt := false;
  TRY
    CNX.Connected := True;
    Cnx.BeginTrans;
    TRY
      //
      QQ := TADOQuery.Create(Application);
      QQ.Connection := CNX;
      if wExistTable ('CPSEQCORRESP') then
      begin
        NomPartage := GetPartage('DESEQUENCES');
        if NomPartage <> '' then NomReel := NomPartage+'.dbo.DESEQUENCES'
                            else NomReel := 'DESEQUENCES';
        NomPartage := GetPartage('CPSEQCORRESP');
        if NomPartage <> '' then NomReel1 := NomPartage+'.dbo.CPSEQCORRESP'
                            else NomReel1 := 'CPSEQCORRESP';
        SQL := 'SELECT DSQ_VALEUR,DSQ_INCREMENT FROM '+NomReel+' WHERE DSQ_CODE=(SELECT CSC_SEQUENCE FROM '+NomReel1+' WHERE CSC_METIER='''+Cle+''')';
      end else
      begin
        NomPartage := GetPartage('DESEQUENCES');
        if NomPartage <> '' then NomReel := NomPartage+'.dbo.DESEQUENCES'
                            else NomReel := 'DESEQUENCES';
        SQL := 'SELECT DSQ_VALEUR,DSQ_INCREMENT FROM '+NomReel+' WHERE DSQ_CODE='''+Cle+'''';
      end;
      QQ.SQL.Text := SQL;
      QQ.Prepared := True;
      QQ.Open;
      TRY
        if not QQ.Eof then
        begin
          OldValue := QQ.Fields[0].AsInteger;
          Result := OldValue + QQ.fields[1].AsInteger;
          QQ.Active := false;
          // -- préparation requete de maj
          if wExistTable ('CPSEQCORRESP') then
          begin
            NomPartage := GetPartage('DESEQUENCES');
            if NomPartage <> '' then NomReel := NomPartage+'.dbo.DESEQUENCES'
                                else NomReel := 'DESEQUENCES';
            NomPartage := GetPartage('CPSEQCORRESP');
            if NomPartage <> '' then NomReel1 := NomPartage+'.dbo.CPSEQCORRESP'
                                else NomReel1 := 'CPSEQCORRESP';
            SQL := 'UPDATE '+NomReel+' SET DSQ_VALEUR='+IntToStr(result)+' WHERE DSQ_CODE=(SELECT CSC_SEQUENCE FROM '+NomReel1+' WHERE CSC_METIER='''+Cle+''') AND DSQ_VALEUR='+InttoStr(OldValue);
          end else
          begin
            NomPartage := GetPartage('DESEQUENCES');
            if NomPartage <> '' then NomReel := NomPartage+'.dbo.DESEQUENCES'
                                else NomReel := 'DESEQUENCES';
            SQL := 'UPDATE '+NomReel+' SET DSQ_VALEUR='+IntToStr(result)+' WHERE DSQ_CODE='''+Cle+''' AND DSQ_VALEUR='+InttoStr(OldValue);
          end;
          //
          QQ.SQL.Clear;
          QQ.sql.text := SQL;
          if QQ.ExecSQL = 0 then
          begin
            Result := -1;
            Raise Exception.Create('Erreur mise à jour compteur pièece comptable');
          end;
        end;
      FINALLY
        QQ.active := False;
        QQ.Free;
      end;
      CNX.CommitTrans;
    EXCEPT
      on E:Exception do
      begin
        Application.MessageBox (PAnsiChar('Erreur '+#10#13+E.message),'') ;
        CNX.RollbackTrans;
        Raise;
      end;
    end;
  finally
    CNX.Close;
    Cnx.Free;
  end;

end;

function CreateSequence(Cle: string; valeur : integer; xx : integer=0; entity : integer=0) : Boolean;

  function FormateCorrespondance (numero : integer) : string;
  begin
    result := stringreplace (Format('%35d', [numero + 1]), ' ', '0', [rfReplaceAll]);
  end;

var TT,TT1 : TOB;
    INumSeq : integer;
    QQ : TQuery;
    Clef : string;
begin
	result := false;
  INumSeq := 0;
  CLef := '';
  if wExistTable ('CPSEQCORRESP') then
  begin
    QQ := OpenSQL('SELECT MAX(CSC_SEQUENCE) AS MAXSEQ FROM CPSEQCORRESP',true,1,'',true);
    if not QQ.eof then
    begin
      INumSeq := QQ.fields[0].AsInteger;
    end;
    ferme (QQ);
    TT1 := TOB.Create('CPSEQCORRESP',nil,-1);
    TT := TOB.Create('DESEQUENCES',nil,-1);
    TRY
      Inc(INumSeq);
      Clef := FormateCorrespondance(INumSeq);
      TRY
        TT1.SetString('CSC_METIER',cle);
        TT1.SetString('CSC_SEQUENCE',Clef);
        TT1.InsertDB(nil);
      EXCEPT
        raise Exception.Create(traduirememoire('Problème en création de sequence '+Cle));
        Exit;
      end;
      //
      TRY
        TT.PutValue('DSQ_CODE',Clef);
        TT.PutValue('DSQ_VALEUR',valeur);
        TT.PutValue('DSQ_INCREMENT',1);
        TT.InsertDB(nil);
        result := true;
      EXCEPT
      END;
    FINALLY
      TT.Free;
      TT1.Free;
    END;
  end else
  begin
    TT := TOB.Create('DESEQUENCES',nil,-1);
    try
      try
        TT.PutValue('DSQ_CODE',Cle);
        TT.PutValue('DSQ_VALEUR',valeur);
        TT.PutValue('DSQ_INCREMENT',1);
        TT.InsertDB(nil);
        result := true;
      except
        raise Exception.Create(traduirememoire('Problème en création de sequence '+Cle));
      end;
    finally
      TT.free;
    end;
  end;
end;

function CreerCompteur(TypeSouche,CodeSouche:string;DD:tDateTime;Compteur:integer;MultiExo:boolean;Entity:integer=0):integer;
var Cle : string;
begin
  Cle:=CleSequence(TypeSouche,CodeSouche,DD,MultiExo);
  if not ExistSequence(Cle) then
    CreateSequence(Cle,Compteur);
  Result:=ReadCurrentSequence(Cle);
end;

function VerifCaractereInterditSouche : string ;
Var ST,NewSt,SauveSt,SH_TYPE : String ;
    Q : tquery ;
    TobS,TobL : tob ;
    TobSAll : tob ;
    i : Integer ;
    Car : Char ;
    OkOk :Boolean ;
begin
  Result:='' ;
  TobS := TOB.Create('', nil, -1);
  TobSAll := TOB.Create('', nil, -1);
  try
    Try
    St:='SELECT * FROM SOUCHE WHERE SH_TYPE IN ("CPT","BUD","REL","TRE") AND SH_SOUCHE LIKE "%.%" ' ;
    Q:=OpenSQL(St,True) ;
    TobS.LoadDetailDB ('SOUCHE', '', '', Q, True);
    Ferme(Q) ;

    St:='SELECT * FROM SOUCHE WHERE SH_TYPE IN ("CPT","BUD","REL","TRE") ' ;
    Q:=OpenSQL(St,True) ;
    TobSAll.LoadDetailDB ('SOUCHE', '', '', Q, True);
    Ferme(Q) ;

    If TobS.Detail.Count>0 then
      begin
      For i:=0 To TobS.Detail.Count-1 Do
        BEGIN

        TobL:=TobS.Detail[i] ;
        SH_TYPE:=TobL.GetValue('SH_TYPE') ;
        St:=TobL.GetValue('SH_SOUCHE') ; SauveSt:=St ;
        Car:='A' ;
        Repeat
        NewSt:=FindEtReplace(St, '.', Car, True);
        OkOk:=TRUE ;

        If TobSAll.FindFirst(['SH_TYPE','SH_SOUCHE'],[SH_TYPE,NewSt],TRUE) <>NIL Then
           begin
           OkOK:=False ;
           If car='Z' then BEGIN OkOk:=TRUE ; Result:=Result+SH_TYPE+';'+SauveSt+';' ; END  Else Car:=Succ(Car) ;
           END ;
        If Not OkOk Then St:=SauveSt ;
        until OkOk ;
        If OkOk then
          begin
          If SH_TYPE='CPT' then
            begin
            ExecuteSQL('UPDATE JOURNAL SET J_COMPTEURNORMAL="'+NewSt+'" WHERE J_COMPTEURNORMAL="'+SauveSt+'" ') ;
            ExecuteSQL('UPDATE JOURNAL SET J_COMPTEURSIMUL="'+NewSt+'" WHERE J_COMPTEURSIMUL="'+SauveSt+'" ');
            end Else
          If SH_TYPE='BUD' then
            begin
            ExecuteSQL('UPDATE BUDJAL SET BJ_COMPTEURNORMAL="'+NewSt+'" WHERE BJ_COMPTEURNORMAL="'+St+'" ');
            ExecuteSQL('UPDATE BUDJAL SET BJ_COMPTEURSIMUL="'+NewSt+'" WHERE BJ_COMPTEURSIMUL="'+St+'" ');
            end ;
          ExecuteSQL('UPDATE SOUCHE SET SH_SOUCHE="'+NewSt+'" WHERE SH_TYPE="'+SH_TYPE+'" AND SH_SOUCHE="'+SauveSt+'" ')
          end;
        END ;
      END ;
    except
      Result:='ERREUR' ;
      raise Exception.Create(traduirememoire('Création Séquence : Problème de transformation des codes souches (Caractère ".")'));
    End ;
  Finally
    TobS.ClearDetail ; TobS.Free ;
    TobSAll.ClearDetail ; TobSAll.Free ;
  end ;
END ;

function OkSouche(Q : tQuery ; St : String) : Boolean ;
Var LeTypeSouche,LaSouche : string ;
begin
  Result:=TRUE ; if St='' Then Exit ;
  LeTypeSouche:=ReadTokenSt(st) ; LaSouche:=ReadTokenSt(st) ;
  If LeTypeSouche='' Then Exit ; If LaSouche='' Then Exit ;
  If (Q.FindField('SH_TYPE').AsString=LeTypeSouche) And (Q.FindField('SH_SOUCHE').AsString=LaSouche) Then Result:=FALSE ;
END ;

Procedure CPInitSequenceCompta;
//Lek 170409 Créer les séquences compta selon table SOUCHE de toutes les entités
var lQ : tQuery;
    req,St : String;
begin
St:=VerifCaractereInterditSouche ;
If St='ERREUR' Then Exit ;
  req:='SELECT * FROM SOUCHE WHERE SH_TYPE IN ("CPT","BUD","REL","TRE")';
  try
    lQ:=OpenSql(req,TRUE);
    while not lQ.Eof do
    begin
      If OkSouche(lQ,St) then
      BEGIN
        if lQ.FindField('SH_SOUCHEEXO').AsString='X' then
        begin
          if lQ.FindField('SH_NUMDEPARTS').AsInteger >1 then
            CreerCompteur(lQ.FindField('SH_TYPE').AsString,
                          lQ.FindField('SH_SOUCHE').AsString,
                          GetSuivant.Deb,
                          lQ.FindField('SH_NUMDEPARTS').AsInteger,
                          lQ.FindField('SH_SOUCHEEXO').AsString='X',
                          lQ.FindField('SH_ENTITY').AsInteger);
          if lQ.FindField('SH_NUMDEPARTP').AsInteger >1 then
            CreerCompteur(lQ.FindField('SH_TYPE').AsString,
                          lQ.FindField('SH_SOUCHE').AsString,
                          GetPrecedent.Deb,
                          lQ.FindField('SH_NUMDEPARTP').AsInteger,
                          lQ.FindField('SH_SOUCHEEXO').AsString='X',
                          lQ.FindField('SH_ENTITY').AsInteger);
        end;
        if lQ.FindField('SH_NUMDEPART').AsInteger >0 then
          CreerCompteur(lQ.FindField('SH_TYPE').AsString,
                        lQ.FindField('SH_SOUCHE').AsString,
                        GetEncours.Deb,
                        lQ.FindField('SH_NUMDEPART').AsInteger,
                        lQ.FindField('SH_SOUCHEEXO').AsString='X',
                        lQ.FindField('SH_ENTITY').AsInteger);
      END ;
      lQ.Next;
    end;
  finally Ferme(lQ); end;
end;

// -------------------------------------------------------

function ConstitueCodeSequenceGC (TypeSouche,CodeSouche : string) : string;
begin
  result := 'SH~'+TypeSouche+'~'+CodeSouche;
end;

procedure CreateSequenceSouche (TypeSouche,CodeSouche: string;  Numero : integer) ;
var CodeSequence : string;
begin
	CodeSequence := ConstitueCodeSequenceGC (TypeSouche,CodeSouche);
  if not ExistSequence(CodeSequence) then
  	CreateSequence (CodeSequence,Numero);
end;

procedure GCInitSequenceSouche;
var
  TobSouche : TOB;
  iSouche : integer;
  CodeSouche,TypeSouche : string;
  Numero : integer;
begin
  TobSouche := Tob.Create ('LES SOUCHES GS', nil, -1);
  try
    TobSouche.LoadDetailFromSQL('SELECT * FROM SOUCHE WHERE SH_TYPE="GES"');
    for iSouche := 0 to TobSouche.Detail.Count - 1 do
    begin
      CodeSouche := TobSouche.Detail[iSouche].GetString ('SH_SOUCHE');
      TypeSouche := TobSouche.Detail[iSouche].GetString ('SH_TYPE');
      Numero := TobSouche.Detail[iSouche].GetInteger('SH_NUMDEPART');
      try
        // CRM_20090924_MNG_FQ;500;16769
        //GCSoucheSetValue (CodeSouche, 0);
        CreateSequenceSouche (TypeSouche,CodeSouche, Numero) ;
      except
      on e : Exception do
        //
      end;
    end;
  finally
    TobSouche.Free;
  end;
end;

end.
