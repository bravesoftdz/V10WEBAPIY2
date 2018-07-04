{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVSELFORMULE ()
Mots clefs ... : TOF;AFREVSELFORMULE
*****************************************************************}
Unit uTofAfRevSelFormule;

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
     UTOF,UTob,HTB97,
     AGLInit;

Type

  TDerniereOuProchaine =(Non,Derniere,Prochaine) ;
  TOF_AFREVSELFORMULE = Class (TOF)
   private
    BTNValider : TToolbarButton97 ;
    Affaire : String ;
    MaGrille : THGrid ;
    iDerniereOuProchaine : TDerniereOuProchaine ;
    Function DerniereOuProchaine(affaire,formule : string) : TDerniereOuProchaine ;
   public
    procedure MaGrilleClick(sender : Tobject) ;
    procedure BTNValiderClick(sender : Tobject) ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  end ;
Const
ColTitre  = 0;
ColDate = 1 ;
ColIcone  =2 ;
ColString  =3 ;
ColCoche =4 ;
StNon='#ICO#68' ;
StDerniere='#ICO#63' ;
STProchaine='#ICO#31' ;
                     
procedure AFLanceFiche_CalcFormule(range,cle,Action : string ) ;

Implementation

procedure TOF_AFREVSELFORMULE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFREVSELFORMULE.OnDelete ;
begin
  Inherited ;
end ;

procedure decodeIcone(St : String ; var PROCHAINE,RECALCUL : String) ;
begin
if st=STProchaine then PROCHAINE:='X' ;
if st=StNon then RECALCUL:='-'  ;
end ;

Function DecodeRecalcul(St : String  ) : String ;
begin
if (St='Oui') or (St='X') then Result:='X' Else Result:='-' ;
end ;


procedure TOF_AFREVSELFORMULE.BTNValiderClick(sender : Tobject) ;
var
  i         : Integer;
  TobFille  : Tob;
  RECALCUL, PROCHAINE : string ;

begin
nextprevcontrol(Ecran) ;
  theTob := TOB.Create('Retour des saisies ',nil,-1);
  for i := 1 to MaGrille.RowCount-1 do
  begin
    TobFille:=TOB.Create('Une Ligne de la grille',theTob,-1);
    RECALCUL := DecodeRecalcul(MaGrille.cells[ColCoche,i]) ;
    PROCHAINE:='-' ;
    decodeIcone(MaGrille.cells[ColIcone,i],PROCHAINE,RECALCUL) ;
    TobFille.AddChampSupValeur('FORMULE',MaGrille.cells[ColTitre,i]) ;
    TobFille.AddChampSupValeur('PROCHAINE',PROCHAINE) ;
    TobFille.AddChampSupValeur('RECALCUL',RECALCUL) ;
  end;
end;

procedure TOF_AFREVSELFORMULE.MaGrilleClick(sender : Tobject) ;
begin

end ;


procedure TOF_AFREVSELFORMULE.OnUpdate ;
begin
  Inherited ;
end ;

Function TOF_AFREVSELFORMULE.DerniereOuProchaine(affaire,formule : string) : TDerniereOuProchaine ;
  Var St : String ;
  Q : TQuery ;
  T : Tob;
begin                            
  T:=TOB.Create('Derniere Ou Prochaine ',nil,-1);
  st:='SELECT MAX(AFR_DATECALCCOEF) AS MADATE, AFR_OKCOEFAPPLIQUE FROM AFREVISION  WHERE AFR_AFFAIRE = "'+affaire+'"' ;
  st:=st+' AND AFR_FORCODE = "'+formule+'" AND AFR_COEFREGUL = "-" GROUP BY AFR_OKCOEFAPPLIQUE ' ;
  Q:=Nil ;
  try
  Q := OpenSQL(St, TRUE);
  T.LoadDetailDB('','','',Q,false) ;
  finally
  Ferme(Q) ;
  end ;
  if (T.Detail.Count > 1) and 
     (T.detail[0].GetValue('MADATE') = T.detail[1].GetValue('MADATE')) then
  begin
  // on a les deux cas
  result:=Non ;
  end else
  begin
  if T.Detail.Count>0 then
    begin
    if T.Detail[0].getvalue('AFR_OKCOEFAPPLIQUE')='X' then
       result:=Prochaine else
       result:=Derniere ; // on peut recalculer
    end else
    begin
    result:=Prochaine ; // prochaine
    end   ;
  end ;
  T.free ;
end ;


procedure TOF_AFREVSELFORMULE.OnLoad ;
Var
  ChampDate   : String;
  St          : String;
  AFC_FORCODE : String;
  Q           : TQuery;
  TobParamF   : tob;
  stMessage   : String;
  StIcone     : String;
  i           : integer;
  
begin

  TobParamF:=TOB.Create('Mes Formules de l''affaire',nil,-1);
  St:='SELECT AFC_PREMIEREDATE,AFC_FORCODE,AFC_AFFAIRE,AFC_LASTDATEAPP,AFC_NEXTDATEAPP, AFC_LASTDATECALC ' ;
  St:=St+' FROM AFPARAMFORMULE WHERE  AFC_AFFAIRE="'+Affaire+'"';
  Q:=Nil ;
  try
  Q := OpenSQL(St, TRUE);
  TobParamF.LoadDetailDB('','','',Q,false) ;
  finally
  Ferme(Q) ;
  end ;


  i:=0  ;

  //Magrille.colTypes[ColCoche]:='B' ;
  //Magrille.colFormats[ColCoche]:=inttostr(integer(csCheckBox)) ;
  Magrille.rowcount:=TobParamF.Detail.count+2 ; // titre+ derniere ligne
  while (i<TobParamF.Detail.count) do
  begin
    AFC_FORCODE:=TobParamF.Detail[i].getvalue('AFC_FORCODE') ;
    iDerniereOuProchaine:=DerniereOuProchaine(affaire,AFC_FORCODE) ;
    if iDerniereOuProchaine=Non then
      begin
      ChampDate:='AFC_LASTDATEAPP' ;
      stMessage:='Application en cours' ;
      StIcone:=StNon ;
      end ;

    if iDerniereOuProchaine=Derniere then
      begin                       
      ChampDate:='AFC_LASTDATECALC' ;
      stMessage:='Recalcul de la révision' ;
      StIcone:=StDerniere ;
      end  ;

    if iDerniereOuProchaine=Prochaine then
       begin
      stMessage:='Calcul de la prochaine révision' ;
      if TobParamF.Detail[i].getvalue('AFC_NEXTDATEAPP')=idate2099 then
        ChampDate:='AFC_PREMIEREDATE'
        Else
        ChampDate:='AFC_NEXTDATEAPP' ;
      StIcone:=STProchaine ;
      end ;

      Magrille.Cells[ColTitre,i+1]:=AFC_FORCODE;
      Magrille.Cells[ColDate,i+1]:=TobParamF.Detail[i].getvalue(ChampDate) ;
      Magrille.Cells[ColString,i+1]:=stMessage ;
      Magrille.Cells[ColIcone,i+1]:=StIcone ;
      Magrille.Cells[ColCoche,i+1]:='Oui' ;
      inc(i) ;
  end ;
  Magrille.rowcount:=Magrille.rowcount-1 ; // j'neleve la derniere ligne
  TobParamF.free ;

end ;

procedure TOF_AFREVSELFORMULE.OnArgument (S : String ) ;
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
        if Champ = 'AFFAIRE' then Affaire := Valeur;
        END;
     Critere:=(Trim(ReadTokenSt(S)));
  END;
  MaGrille:=THGrid(Getcontrol('MaGrille')) ;
  MaGrille.OnClick:=MaGrilleClick ;
  BTNValider:=TToolbarButton97(Getcontrol('BVALIDER')) ;
  BTNValider.OnClick:=BTNValiderClick ;
end ;

procedure TOF_AFREVSELFORMULE.OnClose ;
begin
  Inherited ;
  end ;

procedure TOF_AFREVSELFORMULE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVSELFORMULE.OnCancel () ;
begin
  Inherited ;
end ;

procedure AFLanceFiche_CalcFormule(range,cle,Action : string ) ;
begin
  AglLanceFiche ('AFF','AFREVSELFORMULE',range,cle,Action);
end ;                            


Initialization
  registerclasses ( [ TOF_AFREVSELFORMULE ] ) ;
end.
