{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFUTILREVISION ()
Mots clefs ... : TOF;AFUTILREVISION
*****************************************************************}
Unit uTofAfUtilRevision;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
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
     UTOF,Utob ;

Type
  TOF_AFUTILREVISION = Class (TOF)
    private Formule,Affaire : String ;
      TobF,TobParamF : Tob ;
    public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure   EnleveArrondi(var StMemo : string ) ;
  end ;
procedure AglLanceFicheAFUTILREVISION(cle,Action : string ) ;

Implementation

procedure TOF_AFUTILREVISION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFUTILREVISION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFUTILREVISION.OnUpdate ;
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


procedure   TOF_AFUTILREVISION.EnleveArrondi(var StMemo : string ) ;
var i,PosVirgule,PosParentheseFermee : integer ;
StaJeter,StIndice,StIndiceAff : string ;
begin
  While (pos(',',StMemo)>0) do
    begin
    PosVirgule:=pos(',',StMemo) ;
    PosParentheseFermee:=pos('}',StMemo) ;
    StaJeter:=copy(StMemo,PosVirgule,PosParentheseFermee-PosVirgule+1)  ;
    StMemo:=Stringreplace(StMemo,StaJeter,'', [rfReplaceAll,rfIgnoreCase]) ;
    end ;
  StMemo:=Stringreplace(StMemo,'ARR{','', [rfReplaceAll,rfIgnoreCase]) ;
  For i:=1 to 10 do
    begin
    StIndice:=trim(valeur(TobF.detail[0].getvalue('AFE_INDCODE'+inttostr(i))));
    StIndiceAff:=trim(valeur(TobParamF.detail[0].getvalue('AFC_INDAFF'+inttostr(i))));
    StMemo:=Stringreplace(StMemo,StIndice,StIndiceAff,[rfIgnoreCase]) ;
    end ;
end ;


procedure TOF_AFUTILREVISION.OnLoad ;
  Var St : String ;

  Q : TQuery ;
  StMemo : string ;
begin
  TobF:=TOB.Create('Mes Formules',nil,-1);
  St:='SELECT AFE_FORCODE,AFE_FOREXPRESSION,AFE_INDCODE10, AFE_INDCODE9, AFE_INDCODE8,AFE_INDCODE7, AFE_INDCODE6,' ;
  St:=St+'  AFE_INDCODE5,AFE_INDCODE4,AFE_INDCODE3,AFE_INDCODE2,AFE_INDCODE1 ' ;
  St:=St+' FROM AFFORMULE WHERE AFE_FORCODE = "'+Formule+'"';
  Q:=Nil ;
  try
  Q := OpenSQL(St, TRUE);
  TobF.LoadDetailDB('','','',Q,false) ;
  finally
  Ferme(Q) ;
  end ;

  TobParamF:=TOB.Create('Mes Parametres de Formules',nil,-1);
  St:='SELECT AFC_FORCODE,AFC_AFFAIRE,AFC_INDAFF10, AFC_INDAFF9, AFC_INDAFF8,AFC_INDAFF7, AFC_INDAFF6,' ;
  St:=St+'  AFC_INDAFF5,AFC_INDAFF4,AFC_INDAFF3,AFC_INDAFF2,AFC_INDAFF1 ' ;
  St:=St+' FROM AFPARAMFORMULE WHERE AFC_FORCODE = "'+Formule+'" AND AFC_AFFAIRE="'+Affaire+'"';
  Q:=Nil ;
  try
  Q := OpenSQL(St, TRUE);
  TobParamF.LoadDetailDB('','','',Q,false) ;
  finally
  Ferme(Q) ;
  end ;

  StMemo:=TobF.Detail[0].getvalue('AFE_FOREXPRESSION') ;
  EnleveArrondi(StMemo) ;
  TobF.free ;
  TobParamF.free;
  setControltext('MEMO',StMemo) ;
end ;

procedure TOF_AFUTILREVISION.OnArgument (S : String ) ;
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
        if Champ = 'Formule' then Formule := Valeur;
        if Champ = 'Affaire' then Affaire := Valeur;
        END;
     Critere:=(Trim(ReadTokenSt(S)));
  END;
end ;


procedure TOF_AFUTILREVISION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFUTILREVISION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFUTILREVISION.OnCancel () ;
begin
  Inherited ;
end ;
procedure AglLanceFicheAFUTILREVISION(cle,Action : string ) ;
begin
  AglLanceFiche('AFF','AFUTILREVISION','',cle,Action);
end ;


Initialization
  registerclasses ( [ TOF_AFUTILREVISION ] ) ;
end.
