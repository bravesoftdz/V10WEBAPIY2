unit CalcOLEAffaire;

  //mcd 28/03/03 source revu pour exclure les fct générique GI/GA qui pourront être inclu dans AGL
interface
uses SysUtils, Classes, Utob,
{$IFDEF AFFAIRE}
  AffaireUtil, SyntheseUtil, ParamSoc, utilRevFormule,
{$endif}                                               
  UAFO_Ressource, HEnt1, UtilRessource,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
  CalcOLEGenericBTP,
{$ENDIF}
  HCtrls ,CalcOLEGescom ,DicoBTP ,TraducAffaire,UtilPgi,
  ConfidentAffaire;

function AFCalcOLEEtat(sf,sp : string) : variant ;
{$IFDEF AFFAIRE}
Function ExerN1 (Affaire : string) :string;
{$endif}

implementation             


Function OLEFonctionRessource  (CodeRes : string; Date : TDateTime): variant;
var Res : TAFO_Ressource;
    slFonction : TStringList;
BEGIN
Result := ''; if CodeRes = '' then Exit;
Res := TAFO_Ressource.create (CodeRes);
slFonction := Res.FonctionDeLaRessource(Date);
Result := slFonction[0];
END;


(* mcd 30/01/03 fct identique de CalcOleGescom ==Supprimée
Function NumMois ( Date : TDateTime) : integer;
Var  Annee, Mois, Jour : Word;
BEGIN
DecodeDate(Date, Annee, Mois, Jour);
Result := Mois;
END;  *)

Function LibelleStat (CodeStat : string) : string;
//Var stTablette : string;
BEGIN
Result := ''; if (CodeStat = '') then exit;
Result := RechDomZoneLibre (CodeStat, False);
//Result:=RechDom('GCZONELIBRE',CodeStat,False);
if (result='.-') or(result='.') or (result='-') then result:='';
END;

function AFCalcOLEEtat(sf,sp : string) : variant ;
Var st1,st2,st3,st4 ,st5: string;
    date1,Date2,Date3 : TDateTime;
    //int1    : integer;
{$Ifdef AFFAIRE}
    QQ             : TQuery ;
{$endif}
BEGIN
   //mcd 01/07/2003 AfReadParam a été remplacé par ReadtokenSt ... qui ne fait pas un trim que faisait cette fct !!! fait ..
if sf='CODEAFFUSER' then
   BEGIN
   st1 := trim(ReadTokenSt(sp));
{$IFDEF AFFAIRE}
   Result := CodeAffaireUser(st1);
{$endif}
   Exit;
   END else
if sf='FONCTIONRES' then    // mcd 27/01/03 n'existe plus dans no états modèles
   BEGIN
   st1  := trim(ReadTokenSt(sp)) ;
   Date1:= StrToFloat(trim(ReadTokenSt(sp))) ;
   Result := OLEFonctionRessource(st1,Date1);
   Exit;
   END else
if sf='CONVERSIONUNITE' then
   BEGIN
   st1  := trim(ReadTokenSt(sp)) ;
   st2  := trim(ReadTokenSt(sp)) ;
   st3  := trim(ReadTokenSt(sp)) ;
   if (st1='')or (st2='')or (st3='') then Result:= 0.0 else Result := ConversionUnite(st1, st2, valeur(st3));
   END else
(* mcd 30/01/2003 identique que CalcOleGescom ==> Supprimée
if sf='SEMAINE' then
   BEGIN
   Date1:= StrToFloat(ReadTokenSt(sp)) ;
   Result := NumSemaine (Date1);
   END else
if sf='MOIS' then
   BEGIN
   Date1:= StrToFloat(ReadTokenSt(sp)) ;
   Result := NumMois (Date1);
   END else *)
if sf='ECHEANCEAFFAIRE' then   // mcd 27/01/03 n'existe plus dans no états modèles
   BEGIN
   st1  := trim(ReadTokenSt(sp)) ;
   Date1:= StrToFloat(trim(ReadTokenSt(sp))) ;
{$IFDEF AFFAIRE}
   Result := EcheanceMoisAffaire (st1,Date1);
{$endif}
   END else
if sf='GCECHE' then
   BEGIN
   Result:=GCGetEchesEdt(sf,sp) ;
   END else
if copy(sf,1,6)='GCBASE' then
   BEGIN
   Result:=GCGetBasesEdt(sf,sp) ;
   END else
if sf='TRADUITAF' then    //mcd 31/01/2003 n'existe plus dans nos modèles sauf GPO-GPA qui n'a pas le champs ISGESTFAFF dans le QRS1
    BEGIN
    Result := TraduitGA(sp);
    END else
if sf='TRADUITCOMBO' then  // voir si n'existe pas dans AGL ...
    BEGIN
    Result := TraduitCombo(sp);
    END else
if sf='SYNTHESEAFFRES' then  // mcd 27/01/03 n'existe plus dans no états modèles
    BEGIN
    st1  := trim(ReadTokenSt(sp)) ;    // Code Affaire
    st2  := trim(ReadTokenSt(sp)) ;    // Code Ressource (detail)
    st3  := trim(ReadTokenSt(sp)) ;  //champ cumulé
    st4  := trim(ReadTokenSt(sp)) ;  // Unité
    Date1:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois en cours
    Date2:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois début
    Date3:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois fin
    st5  := sp ;  // Etatvisa  a tjrs mettre en dernier car on prend tout ... même les ;
{$IFDEF AFFAIRE}
    Result := SyntheseAffaireOLE(st1,'ACT_AFFAIRE',st2,'ACT_RESSOURCE',st3,st4,st5, Date1,Date2,Date3);
{$endif}
    if Not(AffichageValorisation) and (St3<>'ACT_QTE') then Result := 0.0;
    END else
if sf='SYNTHESEAFFART' then   // mcd 27/01/03 n'existe plus dans no états modèles
    BEGIN
    st1  := trim(ReadTokenSt(sp)) ;    // Code Affaire
    st2  := trim(ReadTokenSt(sp)) ;    // Code Ressource (detail)
    st3  := trim(ReadTokenSt(sp)) ;  //champ cumulé
    st4  := trim(ReadTokenSt(sp)) ;  // Unité
    Date1:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois en cours
    Date2:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois début
    Date3:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois fin
    st5  := sp ;  // Etatvisa  a tjrs mettre en dernier car on prend tout ... même les ;
{$IFDEF AFFAIRE}
    Result := SyntheseAffaireOLE(st1,'ACT_AFFAIRE',st2,'ACT_ARTICLE',st3,st4,st5, Date1,Date2,Date3);
{$endif}
    if Not(AffichageValorisation)and (St3<>'ACT_QTE') then Result := 0.0;
    END else
if sf='SYNTHESERESAFF' then   // mcd 27/01/03 n'existe plus dans no états modèles
    BEGIN
    st1  := trim(ReadTokenSt(sp)) ;    // Code Affaire
    st2  := trim(ReadTokenSt(sp)) ;    // Code article (detail)
    st3  := trim(ReadTokenSt(sp)) ;  //champ cumulé
    st4  := trim(ReadTokenSt(sp)) ;  // Unité
    Date1:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois en cours
    Date2:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois début
    Date3:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois fin
    st5  := sp ;  // Etatvisa  a tjrs mettre en dernier car on prend tout ... même les ;
{$IFDEF AFFAIRE}
    Result := SyntheseAffaireOLE(st1,'ACT_RESSOURCE',st2,'ACT_AFFAIRE',st3,st4,st5, Date1,Date2,Date3);
{$endif}
    if Not(AffichageValorisation) and (St3<>'ACT_QTE') then Result := 0.0;
    END else
if sf='SYNTHESERESART' then  // mcd 27/01/03 n'existe plus dans no états modèles
    BEGIN
    st1  := trim(ReadTokenSt(sp)) ;    // Code Affaire
    st2  := trim(ReadTokenSt(sp)) ;    // Code Ressource (detail)
    st3  := trim(ReadTokenSt(sp)) ;  //champ cumulé
    st4  := trim(ReadTokenSt(sp)) ;  // Unité
    Date1:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois en cours
    Date2:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois début
    Date3:= StrToFloat(trim(ReadTokenSt(sp))) ;  // Mois fin
    st5  := sp ;  // Etatvisa  a tjrs mettre en dernier car on prend tout ... même les ;
{$IFDEF AFFAIRE}
    Result := SyntheseAffaireOLE(st1,'ACT_RESSOURCE',st2,'ACT_ARTICLE',st3,st4,st5, Date1,Date2,Date3);
{$endif}
    if Not(AffichageValorisation) and (St3<>'ACT_QTE') then Result := 0.0;
    END else
if sf= 'LIBELLESTAT' then   //mcd 31/01/03 n'ets plus utiliser dans nos états modèles. remplacer par traduitcombo
   BEGIN
    st1  := trim(ReadTokenSt(sp)) ;    // Code Stat
    Result := LibelleStat (st1);
   END else
        // mcd 02/10/2002
if sf='AFEXERPREC' then
    BEGIN
    st1  := trim(ReadTokenSt(sp)) ;    // code affaire
{$IFDEF AFFAIRE}
    Result := trim(ExerN1(St1));
{$endif}
    END else
if sf='AFQTEPREC' then
    BEGIN
    st1  := trim(ReadTokenSt(sp)) ;    // code affaire
    st2  := trim(ReadTokenSt(sp)) ;    // code client
{$IFDEF AFFAIRE}
    If GetParamSoc('SO_AFFORMATEXER')='AUC' then Result := 0
    else begin
         st5:='SELECT SUM(ACT_QTEUNITEREF) AS QTE FROM ACTIVITE WHERE ACT_AFFAIRE LIKE "'+st1+'%" AND ACT_TIERS="'+
            st2+'"';
          QQ:=OpenSQL(st5,True,-1,'',true);
          if Not QQ.EOF then
             begin
               Result:=QQ.FindField('QTE').AsFloat ;
               end;
          Ferme(QQ) ;
         end;
{$endif}
    END else
if sf='AFPVPREC' then
    BEGIN
    st1  := trim(ReadTokenSt(sp)) ;    // code affaire
    st2  := trim(ReadTokenSt(sp)) ;    // code client
{$IFDEF AFFAIRE}
    If GetParamSoc('SO_AFFORMATEXER')='AUC' then Result := 0
    else begin
         st5:='SELECT SUM(ACT_TOTVENTE) AS TOTVENTE FROM ACTIVITE WHERE ACT_AFFAIRE LIKE "'+st1+'%" AND ACT_TIERS="'+
            st2+'"';
          QQ:=OpenSQL(st5,True,-1,'',true);
          if Not QQ.EOF then
             begin
               Result:=QQ.FindField('TOTVENTE').AsFloat ;
               end;
          Ferme(QQ) ;
         end;
{$endif}
    END else
if sf='AFFACPREC' then
    BEGIN
    st1  := trim(ReadTokenSt(sp)) ;    // code affaire
    st2  := trim(ReadTokenSt(sp)) ;    // code client
{$IFDEF AFFAIRE}
    If GetParamSoc('SO_AFFORMATEXER')='AUC' then Result := 0
    else begin
         st5:='SELECT SUM(GL_TOTALHT) AS TOTVENTE FROM LIGNE WHERE GL_AFFAIRE LIKE "'+st1+'%" AND GL_TIERS="'+
            st2+'" AND (GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="FRE" OR GL_NATUREPIECEG="AVC") AND GL_TYPELIGNE="ART"';
          QQ:=OpenSQL(st5,True,-1,'',true);
          if Not QQ.EOF then
             begin
               Result:=QQ.FindField('TOTVENTE').AsFloat ;
               end;
          Ferme(QQ) ;
         end;
{$ENDIF}
{$IFDEF AFFAIRE}
  END else
if sf='FORMULEREVISION1' then
  begin
    st1  := trim(ReadTokenSt(sp)) ;    // code affaire
    st2  := trim(ReadTokenSt(sp)) ;    // forcode
    result := FormuleEdition(1, st1, st2);
  END else
if sf='FORMULEREVISION2' then
  begin
    st1  := trim(ReadTokenSt(sp)) ;    // code affaire
    st2  := trim(ReadTokenSt(sp)) ;    // forcode
    result := FormuleEdition(2, st1, st2);
  END else
if sf='FORMULEREVISION3' then
  begin
    st1  := trim(ReadTokenSt(sp)) ;    // code affaire
    st2  := trim(ReadTokenSt(sp)) ;    // forcode
    result := FormuleEdition(3, st1, st2);
{$ENDIF}
  END else
  {$IFDEF BTP}
    Result:=  BTPAFGenCalcOLEEtat(sf,sp ); // mcd 28/03/03
  {$ELSE}
    Result:=  AFGenCalcOLEEtat(sf,sp ); // mcd 28/03/03
  {$ENDIF}

 
END;

{$Ifdef AFFAIRE}
Function ExerN1(Affaire : string) :string;
Var Part0,Part1,Part2,Part3,part4 : string;
BEGIN
Result := '';
{$IFDEF BTP}
BTPCodeAffaireDecoupe(Affaire, part0, Part1, Part2, Part3,Part4,taConsult,False);
{$ELSE}
CodeAffaireDecoupe(Affaire, part0, Part1, Part2, Part3,Part4,taConsult,False);
{$ENDIF}
part3:=''; Part4:=''; // comme on part sur l'exercice précédent, il faut effacer les zones compteurs
Part2 := CodificationNewExercice (Part2,False);
Result :=CodeAffaireRegroupe(Part0, Part1, Part2, part3, Part4, taConsult, True,false,True);
Result:=Copy(Result,1,15)+'  ';  // pour enlever le code avenant ... NB ne focntionnera pas si gestion avenant
END;
{$endif}

end.
