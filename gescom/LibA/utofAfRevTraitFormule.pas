{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 04/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVTRAITFORMULE ()
Mots clefs ... : TOF;AFREVTRAITFORMULE
*****************************************************************}
Unit utofAfRevTraitFormule ;

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
     UTOF, UTob, HTB97,
     AGLInit, Dicobtp, UAFO_REVPRIXCALCULCOEF,
     utilRevision, EntGC;


Type
  TOF_AFREVTRAITFORMULE = Class (TOF)
  public
    procedure BTNValiderClick(sender : Tobject) ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    Retablir : Boolean ;
    fBoAppliquer : Boolean;
    TobParamF: tob ;
    BTNValider : TToolbarButton97 ;
    Affaire : String ;
    MaGrille : THGrid ;

    procedure Appliquer;

  end ;

  Const
    ColTitre  = 0;
    ColDate = 1 ;
    ColCoche =2 ;
           
    TexteMessage: array[1..9] of string 	= (
    {1}  'Les coefficients ont été appliqués.',
    {2}  'Erreur lors de l''application des coefficients.',
    {3}  'Appliquer l''ancien coefficient', // titre d'une colonne de la grid
    {4}  'Erreur lors du chargement de la formule.',
    {5}  'Désappliquer le coefficient',
    {6}  'Pas de coefficients en attente d''application.',
    {7}  'Pas de coefficients à désappliquer.',
    {8}  'L''application des coefficients a été annulée.',                                                       
    {9}  'On ne peut pas annuler l''application du coefficient de la formule %s, car des documents existent pour ce coefficient.'
    ) ;
 
procedure AFLanceFiche_TraitFormule(range,cle,Action : string ) ;
function TestTraitement(pStAffaire : String; pBoAppliquer : Boolean) : Boolean;

Implementation


Function DecodeTraitement(St : String  ) : String ;
begin
if (St='Oui') or (St='X') then Result:='X' Else Result:='-' ;
end ;

procedure TOF_AFREVTRAITFORMULE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFREVTRAITFORMULE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFREVTRAITFORMULE.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/06/2003
Modifié le ... :   /  /
Description .. : On ne propose que les coefficients deja calculés
Mots clefs ... :
*****************************************************************}
procedure TOF_AFREVTRAITFORMULE.OnLoad ;
 Var AFC_DATEAPP,St,AFC_FORCODE : String ;
  Q : TQuery ;
//  StMemo : string ;
//  stMessage,StIcone:String ;
  i : integer ;

begin
  TobParamF:=TOB.Create('Mes Formules de l''affaire',nil,-1);

  if not fBoAppliquer then
  begin
    St:='SELECT DISTINCT AFC_AFFAIRE, AFC_FORCODE, AFC_LASTDATEAPP AS AFC_NEXTDATEAPP ' ;
    St:=St+' FROM AFPARAMFORMULE, AFREVISION WHERE  AFC_AFFAIRE="'+Affaire+'"';
    St:=St+' AND AFR_FORCODE = AFC_FORCODE ';
    St:=St+' AND AFR_AFFAIRE = AFC_AFFAIRE ';
    St:=St+' AND AFR_DATECALCCOEF = AFC_LASTDATEAPP';
    St:=St+' AND AFC_LASTDATEAPP < "' + usdatetime(iDate2099) + '"';
  end
  else
  begin
    St := 'SELECT DISTINCT AFC_AFFAIRE, AFC_FORCODE, AFC_NEXTDATEAPP ' ;
    St := St+' FROM AFPARAMFORMULE, AFREVISION WHERE  AFC_AFFAIRE="'+Affaire+'"';
    St := St+' AND AFR_FORCODE = AFC_FORCODE ';
    St := St+' AND AFR_AFFAIRE = AFC_AFFAIRE ';
    St := St+' AND AFR_DATECALCCOEF = AFC_NEXTDATEAPP';
    St := St+' AND AFC_NEXTDATEAPP < "' + usdatetime(iDate2099) + '"';
  end;

  Q:=Nil;
  try
    Q := OpenSQL(St, TRUE);
    TobParamF.LoadDetailDB('','','',Q,false);
  finally
    Ferme(Q);
  end;

  i:=0;
  Magrille.rowcount:=TobParamF.Detail.count+2 ; // titre+ derniere ligne
  while (i<TobParamF.Detail.count) do
  begin
    AFC_FORCODE := TobParamF.Detail[i].getvalue('AFC_FORCODE') ;
    AFC_DATEAPP := TobParamF.Detail[i].getvalue('AFC_NEXTDATEAPP');
    Magrille.Cells[ColTitre,i+1]:=AFC_FORCODE;
    Magrille.Cells[ColDate,i+1]:= AFC_DATEAPP;
    Magrille.Cells[ColCoche,i+1]:='Oui' ;
    inc(i) ;
  end ;
  Magrille.rowcount:=Magrille.rowcount-1 ; // j'neleve la derniere ligne
end;


procedure TOF_AFREVTRAITFORMULE.OnArgument (S : String ) ;
 Var  Critere, Champ, valeur  : String;
  X : integer ;
  T : HTStrings ;
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
        if Champ = 'ACTION' then
        begin
          Retablir := Valeur = 'RETABLIR';
          if retablir then
            fBoAppliquer := True
          else
            fBoAppliquer := Valeur = 'APPLIQUER';
        end;
    END;
    Critere:=(Trim(ReadTokenSt(S)));
  END;
  MaGrille:=THGrid(Getcontrol('MaGrille')) ;
  BTNValider:=TToolbarButton97(Getcontrol('BVALIDER')) ;
  BTNValider.OnClick:=BTNValiderClick ;
  if retablir then
  begin
    T:=Magrille.titres;
    T[2]:= TexteMessage[3];
    Magrille.UpdateTitres ;
  end
  else if not fBoAppliquer then
  begin
    T:=Magrille.titres;
    T[2]:= TexteMessage[5];
    MaGrille.UpdateTitres;
  end;

end;

procedure TOF_AFREVTRAITFORMULE.BTNValiderClick(sender : Tobject) ;
begin
  Transactions(Appliquer, 1);
end;

procedure TOF_AFREVTRAITFORMULE.Appliquer;
var
  AFR_OKCOEFAPPLIQUE  : string ;
  i                   : integer ;
  stSql               : string ;
  NoMessage           : integer ;
  MaFormule           : TCALCULCOEF;
  vStAppliquer        : String;

begin

  NoMessage := 0;
  nextprevcontrol(Ecran) ;
  try
    for i := 1 to MaGrille.RowCount-1 do
    begin
      AFR_OKCOEFAPPLIQUE:= DecodeTraitement(MaGrille.cells[ColCoche,i]) ;
      if AFR_OKCOEFAPPLIQUE='X' then
      begin
        // ancien coef
        if Retablir then
        begin
          stSql:='Update AFREVISION SET AFR_COEFAPPLIQUE=AFR_DERNIERCOEF ' ;
          stSql:=stSql+' WHERE AFR_AFFAIRE="'+Affaire+'"';
          stSql:=stSql+' AND AFR_FORCODE="'+TobParamF.Detail[i-1].getvalue('AFC_FORCODE')+'"' ;
          executesql(stSql);
        end
        // application du coef calculé
        else
        begin
          if fBoAppliquer then
            vStAppliquer := 'X'
          else
            vStAppliquer := '-';

          stSql:='Update AFREVISION SET AFR_OKCOEFAPPLIQUE= "' + vStAppliquer + '" ' ;
          stSql:=stSql+' WHERE AFR_AFFAIRE="'+Affaire+'"';
          stSql:=stSql+' AND AFR_FORCODE="'+TobParamF.Detail[i-1].getvalue('AFC_FORCODE')+'"' ;
          executesql(stSql);
        end;

        // si le coef appliquer est 0, alors on le force à 1
        stSql :='Update AFREVISION SET AFR_COEFAPPLIQUE = 1 ' ;
        stSql := stSql +' WHERE AFR_AFFAIRE="'+ Affaire + '"';
        stSql := stSql +' AND AFR_FORCODE="'+ TobParamF.Detail[i-1].getvalue('AFC_FORCODE')+'"' ;
        stSql := stSql +' AND AFR_COEFAPPLIQUE= 0' ;
        executesql(stSql);
                       
        // retablir les anciennes dates
        if not fBoAppliquer then
        begin
          if not MajDatesRevisionsApresDesapplication(Affaire, TobParamF.Detail[i-1].getvalue('AFC_FORCODE'), strtodate(Magrille.Cells[ColDate,i])) then
          begin
            NoMessage := 9;
            PGIInfoAF (format(TexteMessage[NoMessage], [TobParamF.Detail[i-1].getvalue('AFC_FORCODE')]),'');
          end
          else             
            NoMessage := 8;
        end
        else
        begin
          // on peut forcer l'application du coefficient calculé
          stSql:='Update AFREVISION SET AFR_COEFAPPLIQUE=AFR_COEFCALC, AFR_OKCOEFAPPLIQUE="X" ' ;
          stSql:=stSql+' WHERE AFR_AFFAIRE="'+Affaire+'"';
          stSql:=stSql+' AND AFR_FORCODE="'+TobParamF.Detail[i-1].getvalue('AFC_FORCODE')+'"' ;
          stSql:=stSql+' AND AFR_APPLIQUECOEF = "X"';
          executesql(stSql);

          MaFormule := TCALCULCOEF.create;
          try
            if not MaFormule.LoadFormule(Affaire, TobParamF.Detail[i-1].getvalue('AFC_FORCODE')) then
              NoMessage := 3
            else if not MaFormule.MajDatesRevisionsApresApplication(True, False, nil) then
              NoMessage := 4;
          finally
            MaFormule.Free;
          end;
        end;
      end;
    end;
    if fBoAppliquer then
      NoMessage := 1
    else if NoMessage = 0 then
      NoMessage := 7;
  except
    NoMessage := 2;
  end;
  if NoMessage <> 9 then
    PGIInfoAF (TexteMessage[NoMessage],'');
end;

procedure TOF_AFREVTRAITFORMULE.OnClose ;
begin
  Inherited ;
  TobParamF.free ;
end ;

procedure TOF_AFREVTRAITFORMULE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVTRAITFORMULE.OnCancel () ;
begin
  Inherited ;
end ;

procedure AFLanceFiche_TraitFormule(range,cle,Action : string ) ;
begin
  AglLanceFiche ('AFF','AFREVTRAITFORMULE',range,cle,Action);
end;

function TestTraitement(pStAffaire : String; pBoAppliquer : Boolean) : Boolean;
var
  St   : String;
  Q    : TQuery;
  vTob : Tob;

begin

  result := true;
  if not pBoAppliquer then
  begin
    St:='SELECT COUNT(*) AS NB ' ;
    St:=St+' FROM AFPARAMFORMULE, AFREVISION WHERE  AFC_AFFAIRE="'+ pStAffaire+'"';
    St:=St+' AND AFR_FORCODE = AFC_FORCODE ';
    St:=St+' AND AFR_AFFAIRE = AFC_AFFAIRE ';
    St:=St+' AND AFR_DATECALCCOEF = AFC_LASTDATEAPP';
    St:=St+' AND AFC_LASTDATEAPP < "' + usdatetime(iDate2099) + '"';
  end
  else
  begin
    St:='SELECT COUNT(*) AS NB ' ;
    St:=St+' FROM AFPARAMFORMULE, AFREVISION WHERE  AFC_AFFAIRE="'+ pStAffaire+'"';
    St:=St+' AND AFR_FORCODE = AFC_FORCODE ';
    St:=St+' AND AFR_AFFAIRE = AFC_AFFAIRE ';
    St:=St+' AND AFR_DATECALCCOEF = AFC_NEXTDATEAPP';
    St:=St+' AND AFC_NEXTDATEAPP < "' + usdatetime(iDate2099) + '"';
  end;

  vTob := TOB.Create('mon Test', nil, -1);
  Q:=Nil;
  try
    Q := OpenSQL(St, TRUE);
    vTob.LoadDetailDB('','','',Q,false) ;
    if vTob.detail[0].getvalue('Nb') = 0 then
    begin
      result := false;
      if pBoAppliquer then
        PGIInfoAF (TexteMessage[6],'')
      else
        PGIInfoAF (TexteMessage[7],'');
    end;

  finally
    Ferme(Q);
    vTob.Free;
  end;
end;

Initialization
  registerclasses ( [ TOF_AFREVTRAITFORMULE ] ) ;
end.
