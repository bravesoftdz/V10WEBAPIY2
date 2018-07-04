{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVISIONINDICE ()
Mots clefs ... : TOF;AFREVISIONINDICE
*****************************************************************}
Unit utofAfRevisionIndice;

Interface

Uses StdCtrls,
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$Else}
     MainEagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     UTOF,utob ;

Type
  TOF_AFREVISIONINDICE = Class (TOF)
    private
    TobF : Tob ;
    AFR_FORCODE,AFR_AFFAIRE,AFR_NUMEROLIGNE : String ;
    public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

procedure AglLanceFicheAFREVISIONINDICE(cle,Action : string ) ;

Implementation

procedure TOF_AFREVISIONINDICE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFREVISIONINDICE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFREVISIONINDICE.OnUpdate ;
begin
  Inherited ;
end ;

function valeur(v : variant) : string ;
begin
 Try
 result:=string(v) ;
 except
 result:='' ;
 end ;
end ;



procedure TOF_AFREVISIONINDICE.OnLoad ;
  Var St : String ;
  Q : TQuery ;
  i : integer ;
begin
  TobF:=TOB.Create('Mes indices et leurs valeurs',nil,-1);
  St:='SELECT '  ;
  for i:=1 to 10 do
    begin
    St:=St+' AFE_INDCODE'+inttostr(i)+',AFR_VALINDICE'+inttostr(i)+',' ;
    end ;
  St:=St+' AFE_FORCODE ';
  St:=St+' FROM AFREVISION INNER JOIN AFFORMULE ON AFE_FORCODE=AFR_FORCODE ' ;
  St:=St+' WHERE AFE_FORCODE = "'+AFR_FORCODE+'"';
  St:=St+' AND   AFR_AFFAIRE = "'+AFR_AFFAIRE+'"';
  St:=St+' AND   AFR_NUMEROLIGNE = '+AFR_NUMEROLIGNE ;
  Q:=Nil ;
  try
  Q := OpenSQL(St, TRUE);
  TobF.LoadDetailDB('','','',Q,false) ;
  finally
  Ferme(Q) ;
  end ;
  if tobF.Detail.count>0 then
    begin
    for i:=1 to 10 do
      begin
      setControltext('AFE_INDCODE'+inttostr(i),valeur(TobF.Detail[0].getvalue('AFE_INDCODE'+inttostr(i)))) ;
      setControltext('AFR_VALINDICE'+inttostr(i),valeur(TobF.Detail[0].getvalue('AFR_VALINDICE'+inttostr(i)))) ;
      end ;
   end ;
 end ;

procedure TOF_AFREVISIONINDICE.OnArgument (S : String ) ;

 Var  Critere, Champ, valeur  : String;
  X : integer ;
begin
  Inherited ;
  Critere:=(Trim(ReadTokenSt(S)));
  While (Critere <>'') do
    Begin
    if Critere<>'' then
      Begin
      X:=pos('=',Critere);
      if x<>0 then
         begin
         Champ:=copy(Critere,1,X-1);
         Valeur:=Copy (Critere,X+1,length(Critere)-X);
         end;
      if Champ = 'AFR_FORCODE' then AFR_FORCODE := Valeur;
      if Champ = 'AFR_AFFAIRE' then AFR_AFFAIRE := Valeur;
      if Champ = 'AFR_NUMEROLIGNE' then AFR_NUMEROLIGNE := Valeur;
      END;
    Critere:=(Trim(ReadTokenSt(S)));
    END;
END;

procedure TOF_AFREVISIONINDICE.OnClose ;
begin
  Inherited ;
  TobF.free ;
end ;

procedure TOF_AFREVISIONINDICE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVISIONINDICE.OnCancel () ;
begin
  Inherited ;
end ;


procedure AglLanceFicheAFREVISIONINDICE(cle,Action : string ) ;
begin
  AglLanceFiche('AFF','AFREVISIONINDICE','',cle,Action);
end ;


Initialization
  registerclasses ( [ TOF_AFREVISIONINDICE ] ) ; 
end.
