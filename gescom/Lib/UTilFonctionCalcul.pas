unit UTilFonctionCalcul;

interface

Uses StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,HTB97,ExtCtrls,
     UTOF,Graphics,UTOB,AGLInit,Math ;

Type TabDblChaine = array[0..1] of string;

Type R_FonctCal = RECORD
     Resultat : Double;
     Formule : String;
     Expression : String;
     Affichage : String ;
     Annule : Boolean;
     VarLibelle : array[0..30] of string;
     VarValeur : array[0..30] of Double;
     Affichable : array[0..30] of Boolean;
     Modifiable : array[0..30] of Boolean;
     END;


  Function  AdapteVariable(St : string) : string;
  function  VerifieCrochets(Chaine : string; ModeSilence : boolean=False) : boolean;
  function  VerifieParentheses(Chaine : string; ModeSilence : boolean=False) : boolean;
  function  ReplaceCaractereAccentue(St: string): string ;
  function  SupprimeCaracteresSpeciaux(St: string ; Accent,AltGr,Operateur : boolean; Chiffresetlettresuniquement : boolean=false): string ; Overload;
  function  SupprimeCaracteresSpeciaux(var St : String ; CaractereSubstitution : String; PourFichier : Boolean=False;PourArticles : Boolean=false) : Boolean; Overload

  Function  DechiffrageFormule(Formule : string; TOBF,TOBFTab : Tob) : boolean;
  Function  initRecord : R_FonctCal;

  Function  GFormuleVersFormatPolonais(Chaine : string) : string;
  Function  Saisie_FonctionParam(Formule : string) : R_FonctCal;
  Function  RecupVariableFormule(Formule : string) : R_FonctCal;
  Function  EvaluationFormule(Titre,Formule : string) : R_FonctCal; overload;
  Function  EvaluationFormule(Titre : string; RFonct : R_FonctCal; ModeSilence : boolean=False) : R_FonctCal; overload;
  function RemplaceCaracteresSpeciaux(var St : String; CaractereSubstitution : String;Accent,AtlGR,Operateur : Boolean) : String;

const
	// libellés des messages
	Messages : array[0..1] of string = (
          {0}         'Il y a une erreur dans les parenthèses'
          {1}        ,'Il y a une erreur dans les crochets'
                     );

implementation
uses GcFonctionVar_Tof, StrUtils;


{==============================================================================================}
{=================================== Utilitaires d'appel=======================================}
{==============================================================================================}
Function AdapteVariable(St : string) : string;
begin
St:=StringReplace(St,' ','_',[rfReplaceAll]);
St:=StringReplace(St,';','',[rfReplaceAll]);
St:=VireParenthese(St);
St:=StringReplace(St,',','.',[rfReplaceAll]);
St:=ReplaceCaractereAccentue(St);
St:=SupprimeCaracteresSpeciaux(St,False,True,True);
Result:=St;
end;

Function VerifieBrackets(Chaine, BracketOuv,Bracketfer : string) : boolean;
var nbOuvPar, nbFerPar, i_ind1 : integer;
begin
Result := True;
nbOuvPar := 0;
nbFerPar := 0;
i_ind1 := 1;
while i_ind1 <= Length(Chaine) do
  begin
  if Chaine[i_ind1] = BracketOuv then
      Inc(nbOuvPar)
  else if Chaine[i_ind1] = Bracketfer then
      Inc(nbFerPar);
  Inc(i_ind1);
  end;
if nbOuvPar <> nbFerPar then
  begin
  Result := False;
  end;
end;

Function VerifieCrochets(Chaine : string; ModeSilence : boolean=False) : boolean;
begin
Result := VerifieBrackets(Chaine,'[',']');
if Not Result then if Not ModeSilence then PgiBox(Messages[1]);
end;

Function VerifieParentheses(Chaine : string; ModeSilence : boolean=False) : boolean;
begin
Result := VerifieBrackets(Chaine,'(',')');
if Not Result then if Not ModeSilence then PgiBox(Messages[0]);
end;

function  ReplaceCaractereAccentue(St: string): string ;
begin
St:=StringReplace(St,'à','a',[rfReplaceAll]);
St:=StringReplace(St,'â','a',[rfReplaceAll]);
St:=StringReplace(St,'ç','c',[rfReplaceAll]);
St:=StringReplace(St,'é','e',[rfReplaceAll]);
St:=StringReplace(St,'è','e',[rfReplaceAll]);
St:=StringReplace(St,'ê','e',[rfReplaceAll]);
St:=StringReplace(St,'ë','e',[rfReplaceAll]);
St:=StringReplace(St,'î','i',[rfReplaceAll]);
St:=StringReplace(St,'ï','i',[rfReplaceAll]);
St:=StringReplace(St,'ô','o',[rfReplaceAll]);
St:=StringReplace(St,'ù','u',[rfReplaceAll]);
St:=StringReplace(St,'û','u',[rfReplaceAll]);
Result:=St;
end;

function RemplaceCaracteresSpeciaux(var St : String; CaractereSubstitution : String;Accent,AtlGR,Operateur : Boolean) : String;
Const Car1 = 'àâäçéèêëîïôöùû';
      Car2 = '&$#@§€"\?,.;:!%£¨°|{}[]<>^¤';
      Car3 = '+/*=';
Var   NbCar         : Integer;
      Caractere     : String;
begin

  NbCar := 1;    
  if (Accent) then
  begin
    While NbCar < length(Car1) Do
    begin
      Caractere := Car1[Nbcar];
      if Pos(Caractere, st) <> 0 then
      begin
        st := AnsiReplaceStr(st, Caractere, CaractereSubstitution);
      end;
      inc(NbCar);
    end;
  end;

  NbCar := 1;
  if (AtlGR) then
  begin
    While NbCar < length(Car2) Do
    begin
      Caractere := Car2[Nbcar];
      if Pos(Caractere, st) <> 0 then
      begin
        st := AnsiReplaceStr(st, Caractere, CaractereSubstitution);
      end;
      inc(NbCar);
    end;
  end;

  NbCar := 1;
  if (Operateur) then
  begin
    While NbCar < length(Car3) Do
    begin
      Caractere := Car3[Nbcar];
      if Pos(Caractere, st) <> 0 then
      begin
        st := AnsiReplaceStr(st, Caractere, CaractereSubstitution);
      end;
      inc(NbCar);
    end;
  end;
                              
  Result := St;

end;

function  SupprimeCaracteresSpeciaux(St: string ; Accent,AltGr,Operateur : boolean; Chiffresetlettresuniquement : boolean=false): string ;
Const Car1 = ['à','â','ç','é','è','ê','ë','î','ï','ô','ù','û'] ;
      Car2 = ['&','$','#','@','§','€','"','\','|','{','}','[',']','<','>'] ;
      Car3 = ['+','/','*','=','-'] ;
      Car4 = ['0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
Var ind : Integer ;
begin
  Ind := 1;
  while (ind <= Length(St))
  do
  begin
    if (Chiffresetlettresuniquement) then // Dans ce cas on ne supprime pas le caractère mais on le remplace par un zéro
    begin
      if not (St[ind] in Car4) then St[ind]:='0';
      ind := ind + 1;
    end else
    begin
      if (Accent and (St[ind] in Car1)) then Delete(St,ind,1) else ind := ind + 1 ;
      if (AltGr and (St[ind] in Car2)) then Delete(St,ind,1) else ind := ind + 1 ;
      if (Operateur and (St[ind] in Car3)) then Delete(St,ind,1) else ind := ind + 1;
  end;

  end;

  Result:=St;
end;

//maVar = StripSpecialChars(maVar, , "&#{}[]()-|_@°=+*%!?,.;/:'""", "")
function SupprimeCaracteresSpeciaux(var St : String ; CaractereSubstitution : String; PourFichier : Boolean=False;PourArticles : Boolean=false) : Boolean;
const CaractereSpecialPourFichier = 'àâçéèêëîïôùû&$£€#~µ%@§"\|{}[]+=*/-<>°!?,.;:';
const CaractereSpecialPourCode = 'àâçéèêëîïôùû&$£€#~µ%_@§"\|{}[]+=/<>°!?,;:';
const CaractereSpecialPourCodeArt = 'àâçéèêëîïôùû&$£€#~µ%@§"\|{}[]+=/<>°!?,;:';
Var NbCar         : Integer;
    Caractere,CaractereSpecial   : String;

begin

  result := false;

  if PourFichier Then
  begin
    CaractereSpecial := CaractereSpecialPourFichier
  end else if PourArticles then
  begin
    CaractereSpecial := CaractereSpecialPourCodeArt;
  end else
  begin
    CaractereSpecial := CaractereSpecialPourCode;
  end;

  NbCar := 1;

  While NbCar < length(CaractereSpecial) Do
  begin
    Caractere := CaractereSpecial[Nbcar];
    if Pos(Caractere, st) <> 0 then
    begin
      result := true;
      st := AnsiReplaceStr(st, Caractere, CaractereSubstitution);
    end;
    inc(NbCar);
  end;

end;

Function initRecord : R_FonctCal;
Var ind : integer;
begin
Result.Resultat:=0;
Result.Formule:='';
Result.Expression:='';
Result.Annule:=False;
for ind:=0 to High(Result.VarLibelle) do
  begin
  Result.VarLibelle[ind]:='';
  Result.VarValeur[ind]:=0;
  Result.Affichable[ind]:=True;
  Result.Modifiable[ind]:=True;
  end;
end;

{==============================================================================================}
{=================================== Fonctions internes =======================================}
{==============================================================================================}
procedure ChercherNombre(Chaine : string; TOBF,TOBFTab : Tob);
var nbOuvPar, nbFerPar, ind, NbCour : integer ;
    Nombre : string ;
    TOBN : TOB;
begin
nbOuvPar := 0;
nbFerPar := 0;
ind := 1; NbCour:=0;
while ind <= Length(Chaine) do
  begin
  if Chaine[ind] = '(' then
    Inc(nbOuvPar)
  else if Chaine[ind] = ')' then
    Inc(nbFerPar)
  else if (Chaine[ind] = ';') and (nbOuvPar-nbFerPar=0) then
    begin
    Nombre:=Copy(Chaine,1,ind-1);
    Inc(NbCour);
    TOBN:=Tob.Create('_Nombre'+IntToStr(NbCour),TobF,-1);
    TOBN.AddChampSup('TYPE',False);
    TOBN.AddChampSup('VALEUR',False);
    if isnumeric(Nombre,True) then TOBN.PutValue('TYPE','N')else TOBN.PutValue('TYPE','V');
    TOBN.PutValue('VALEUR',Nombre);
    if DechiffrageFormule(Nombre,TOBN,TOBFTab) then TOBN.PutValue('TYPE','F');
    Chaine:=Copy(Chaine,ind+1,Length(Chaine)-ind);
    ind := 0;
    end;
  Inc(ind);
  end;
if Chaine<>'' then
  begin
  Inc(NbCour);
  TOBN:=Tob.Create('_Nombre'+IntToStr(NbCour),TobF,-1);
  TOBN.AddChampSup('TYPE',False);
  TOBN.AddChampSup('VALEUR',False);
  if isnumeric(Chaine,True) then TOBN.PutValue('TYPE','N')else TOBN.PutValue('TYPE','V');
  TOBN.PutValue('VALEUR',Chaine);
  if DechiffrageFormule(Chaine,TOBN,TOBFTab) then TOBN.PutValue('TYPE','F');
  end;
end;

Function DechiffrageFormule(Formule : string; TOBF,TOBFTab : Tob) : boolean;
var ind : integer ;
    Chaine,Fonction : string ;
    TobFCour,TOBT : TOB;
begin
Result:=False;
if Formule='' then exit;
Chaine:=Formule;
ind:=pos('(',Chaine);
if ind=0 then exit;
Fonction:=Copy(Chaine,1,ind-1);
TOBT:=TobFTab.FindFirst(['CO_LIBELLE'],[Fonction],False);
if TOBT=Nil then exit;
if Copy(TOBF.NomTable,1,7)='_Nombre' then TobFCour:=Tob.Create('COMMUN',TOBF,-1)else TobFCour:=TOBF ;
TobFCour.Dupliquer(TobT,False,True);
Chaine:=Copy(Chaine,ind+1,Length(Chaine)-ind-1);
if Chaine='' then exit;
ChercherNombre(Chaine,TobFCour,TOBFTab);
Result:=True;
end;

function ListerVariable(TOBF : TOB) : string;
var TypeF,ListeVar,St : string;
    TOBN : TOB;
    ind : integer ;
begin
ListeVar:='';
For ind:=0 to TOBF.Detail.Count-1 do
  begin
  TOBN:=TOBF.Detail[ind];
  if Not TOBN.FieldExists('NOMBRE') then
    begin
    TOBN.AddChampSup('NOMBRE',False);
    TOBN.PutValue('NOMBRE', '0');
    end;
  TypeF:=TOBN.GetValue('TYPE');
  if TypeF='N' then Continue
  else if TypeF='F' then
    begin
    if TOBN.Detail.Count>0 then
      begin
      ListeVar:=ListeVar+ListerVariable(TOBN.Detail[0]);
      end else Continue;
    end
  else if TypeF='V' then
    begin
    St:=TOBN.GetValue('VALEUR')+';';
    if Pos(St,ListeVar)=0 then ListeVar:=ListeVar+St;
    end;
  end;
Result:=ListeVar;
end;

function RenseignerVariable(Titre,Formule : string; TOBF : TOB; TOBVal : TOB; ModeSilence : Boolean) : TabDblChaine ;
var ListSt,ValSt,StList,StVal,Affichage : string ;
    ListeVar,ValeurVar,ModifVar,VisibleVar : string;
    TOBN,TOBTmp : TOB;
begin
ListeVar:=ListerVariable(TOBF);

StList:=ListeVar; 
if TOBVal<>Nil then Affichage:=TOBVal.GetValue('AFFICHAGE') else Affichage:=Formule;

ListeVar:=''; ValeurVar:=''; ModifVar:='';
While StList<>'' do
  begin
  ListSt:=ReadTokenSt(StList);
  if TOBVal<>Nil then
    begin
    TOBN:=TOBVal.FindFirst(['LIBELLE'],[ListSt],True);
    if TOBN<>Nil then
      begin
      ListeVar:=ListeVar+TOBN.GetValue('LIBELLE')+';';
      ValeurVar:=ValeurVar+FloatToStr(TOBN.GetValue('VALEUR'))+';';
      ModifVar:=ModifVar+TOBN.GetValue('MODIFIABLE')+';';
      VisibleVar:=VisibleVar+TOBN.GetValue('AFFICHABLE')+';';
      end else
      begin
      ListeVar:=ListeVar+ListSt+';';
      ValeurVar:=ValeurVar+'0;';
      ModifVar:=ModifVar+'X;';
      VisibleVar:=VisibleVar+'X;';
      end;
    end else
    begin
    ListeVar:=ListeVar+ListSt+';';
    ValeurVar:=ValeurVar+'0;';
    ModifVar:=ModifVar+'X;';
    VisibleVar:=VisibleVar+'X;';
    end;
  end;
if (ListeVar<>'') and (Not ModeSilence) then
  begin
  TOBTmp:=TOB.Create('Paasage param',Nil,-1);
  TOBTmp.AddChampSupValeur('FORMULE',Formule);
  TOBTmp.AddChampSupValeur('AFFICHAGE',Affichage);
  TOBTmp.AddChampSupValeur('LISTEVAR',ListeVar);
  TOBTmp.AddChampSupValeur('VALEURVAR',ValeurVar);
  TOBTmp.AddChampSupValeur('MODIFVAR',ModifVar);
  TOBTmp.AddChampSupValeur('VISIBLEVAR',VisibleVar);
  TheTOB:=TOBTmp;
  ValeurVar:=GCLanceFiche_FonctionVar('GC','GCFONCTIONVAR','','',Titre);
  TOBTmp.Free;
  end;
StVal:=ValeurVar;
StList:=ListeVar;
Result[0]:=StList;
Result[1]:=StVal;
While StList<>'' do
  begin
  ListSt:=ReadTokenSt(StList);
  ValSt:=ReadTokenSt(StVal);
  TOBN:=TOBF.FindFirst(['TYPE','VALEUR'],['V',ListSt],True);
  while TOBN<>Nil do
    begin
    TOBN.PutValue('NOMBRE', ValSt);
    TOBN:=TOBF.FindNext(['TYPE','VALEUR'],['V',ListSt],True);
    end;
  end;
end;

Function CalculFormule(TOBF : TOB) : R_FonctCal;
var TypeF,Fonct,Expression,ExpDeb,ExpSepa,ExpFin : string;
    Symbole,SymbDeb,SymbSepa,SymbFin,St : string;
    TOBN : TOB;
    Resultat : Double ;
    ind : integer ;
    ResCalc : R_FonctCal;
begin
Resultat:=0; Expression:='(';
Fonct:=TOBF.GetValue('CO_CODE');
St:=TOBF.GetValue('CO_ABREGE');
ReadTokenSt(St); ReadTokenSt(St);
Symbole:=ReadTokenSt(St);
If Pos('(',Symbole)>0 then
   begin
   SymbDeb:=Symbole ;
   SymbSepa:=',';
   SymbFin:=')' ;
   end else
   begin
   SymbDeb:='' ;
   SymbSepa:=Symbole;
   SymbFin:='' ;
   end;
For ind:=0 to TOBF.Detail.Count-1 do
  begin
  if ind=0 then
    begin
    ExpDeb:=SymbDeb;
    ExpSepa:='';
    ExpFin:='';
    end else if ind=TOBF.Detail.Count-1 then
    begin
    ExpDeb:='';
    ExpSepa:=' '+SymbSepa+' ';
    ExpFin:=SymbFin;
    end else
    begin
    ExpDeb:='';
    ExpSepa:=' '+SymbSepa+' ';
    ExpFin:='';
    end;
  TOBN:=TOBF.Detail[ind];
  if Not TOBN.FieldExists('NOMBRE') then TOBN.AddChampSup('NOMBRE',False);
  TypeF:=TOBN.GetValue('TYPE');
  if TypeF='N' then
    begin
    TOBN.PutValue('NOMBRE',TOBN.GetValue('VALEUR'));
    Expression:=Expression+ExpDeb+ExpSepa+TOBN.GetValue('NOMBRE')+ExpFin;
    end
  else if TypeF='F' then
    begin
    if TOBN.Detail.Count>0 then
      begin
      ResCalc:=CalculFormule(TOBN.Detail[0]);
      TOBN.PutValue('NOMBRE',floatToStr(ResCalc.Resultat));
      end else
      begin
      ResCalc.Resultat:=0;ResCalc.Expression:='0';
      TOBN.PutValue('NOMBRE','0');
      end;
    Expression:=Expression+ExpDeb+ExpSepa+ResCalc.Expression+ExpFin;
    end
  else if TypeF='V' then
    begin
    //Expression:=Expression+ExpDeb+ExpSepa+TOBN.GetValue('VALEUR')+': '+TOBN.GetValue('NOMBRE')+ExpFin;
    Expression:=Expression+ExpDeb+ExpSepa+TOBN.GetValue('NOMBRE')+' '+TOBN.GetValue('VALEUR')+ExpFin;
    end;
  if Fonct='C00' then // Arrondi
    begin
    if ind=0 then Resultat:= Valeur(TOBN.GetValue('NOMBRE'))
             else Resultat:=Arrondi(Resultat,StrToInt(TOBN.GetValue('NOMBRE')));
    end
  else if Fonct='C10' then // Division
    begin
     if ind=0 then Resultat:= Valeur(TOBN.GetValue('NOMBRE'))
              else if Valeur(TOBN.GetValue('NOMBRE'))<>0 then Resultat:=Resultat/Valeur(TOBN.GetValue('NOMBRE'));
    end
  else if Fonct='C20' then // Ent
    Resultat:= Arrondi(Valeur(TOBN.GetValue('NOMBRE'))-0.499,0)
  else if Fonct='C30' then // Max
    begin
    if ind=0 then Resultat:= Valeur(TOBN.GetValue('NOMBRE'))
             else Resultat:=Max(Resultat,Valeur(TOBN.GetValue('NOMBRE')));
    end
  else if Fonct='C40' then // Min
    begin
    if ind=0 then Resultat:= Valeur(TOBN.GetValue('NOMBRE'))
             else Resultat:=Min(Resultat,Valeur(TOBN.GetValue('NOMBRE')));
    end
  else if Fonct='C50' then // Produit
    if ind=0 then Resultat:= Valeur(TOBN.GetValue('NOMBRE'))
             else Resultat:=Resultat* Valeur(TOBN.GetValue('NOMBRE'))
  else if Fonct='C60' then // Somme
    Resultat:=Resultat+Valeur(TOBN.GetValue('NOMBRE'))
  else if Fonct='C70' then // Soustraction
    if ind=0 then Resultat:= Valeur(TOBN.GetValue('NOMBRE'))
             else Resultat:=Resultat-Valeur(TOBN.GetValue('NOMBRE'));
  end;
//Expression:=Copy(Expression,1,Length(Expression)-1)+')';
Expression:=Expression+')';
Result.Resultat:=Resultat;
Result.Expression:=Expression;
end;

Function EvaluerFormule(Titre,Formule : string; TOBVal : TOB; ModeSilence : Boolean) : R_FonctCal;
var TOBF,TOBFTab,TobN : Tob ;
    QQ : TQuery ;
    ListeVar : TabDblChaine;
    ind : integer;
begin
Result:=initRecord;
if Formule='' then exit;
TobF:=Tob.Create('COMMUN',Nil,-1);
TobFTab:=Tob.Create('',Nil,-1);
QQ:=OpenSql('Select * from Commun Where CO_TYPE="GFX"',True);
if not QQ.Eof then TobFTab.LoadDetailDB('COMMUN','','',QQ,False);
Ferme(QQ);
if DechiffrageFormule(Formule,TobF,TobFTab) then
  begin
  ListeVar:=RenseignerVariable(Titre,Formule,TOBF,TOBVal,ModeSilence);
  Result:=CalculFormule(TOBF);
  ind:=0;
  if ListeVar[1]='' then Result.Annule:=True;
  While (ListeVar[0]<>'') and (ind < High(Result.VarLibelle)) do
    begin
    Result.VarLibelle[ind]:=ReadTokenSt(ListeVar[0]);
    Inc(ind);
    end;
  ind:=0;
  While (ListeVar[1]<>'') and (ind < High(Result.VarValeur)) do
    begin
    Result.VarValeur[ind]:=Valeur(ReadTokenSt(ListeVar[1]));
    Inc(ind);
    end;
  end;
Result.Formule:=Formule;
if TOBVal<>Nil then
  begin
  for ind:=0 to Min(High(Result.VarLibelle),TOBVal.Detail.count-1) do
    begin
    TOBN:=TOBVal.FindFirst(['LIBELLE'],[Result.VarLibelle[ind]],True);
    Result.Affichable[ind]:=(TOBN.GetValue('AFFICHABLE')='X');
    Result.Modifiable[ind]:=(TOBN.GetValue('MODIFIABLE')='X');
    end;
  end;
TobF.Free; TobFTab.Free;
end;

Function  ExtraireSousChaine(Chaine : string) : string;
var nbOuvPar, i_ind1 : integer;
begin
Result := '';
nbOuvPar := 1;
i_ind1 := 1;
while (i_ind1 <= Length(Chaine)) and (nbOuvPar<>0) do
  begin
  if Chaine[i_ind1] = '(' then
      Inc(nbOuvPar)
  else if Chaine[i_ind1] = ')' then
      Dec(nbOuvPar);
  Inc(i_ind1);
  end;
if nbOuvPar=0 then
  begin
  Result := Copy(Chaine,1,i_ind1-2);
  end;
end;

Function  ReStructureGformule(Chaine : string) : string;
Const Fonct = ['+','-','*','/'] ;
Var i_ind1,OuvPar,FerPar : integer;
    SousChaine,FonctPrec,FonctCour,NewChaine : string ;
begin
Result:='';
if Chaine='' then exit;
i_ind1:=1;
FonctPrec:=''; FonctCour:=''; NewChaine:='';
while i_ind1 <= Length(Chaine) do
  begin
  if Chaine[i_ind1] = '(' then
    begin
    OuvPar:=i_ind1;
    SousChaine:=ExtraireSousChaine(Copy(Chaine,i_ind1+1,Length(Chaine)));
    FerPar:=Length(SousChaine);
    SousChaine:=ReStructureGformule(SousChaine);
    i_ind1:=OuvPar+Ferpar+1;
    NewChaine:=NewChaine+'('+SousChaine+')';
    end
  else if Chaine[i_ind1] in Fonct then
    begin
    if FonctPrec='' then FonctPrec:=Chaine[i_ind1];
    FonctCour:=Chaine[i_ind1];
    if FonctPrec<>FonctCour then
      begin
      NewChaine:='('+NewChaine+')'+Chaine[i_ind1];
      FonctPrec:=FonctCour;
      end else
      begin
      NewChaine:=NewChaine+Chaine[i_ind1];
      end;
    end else NewChaine:=NewChaine+Chaine[i_ind1];
    Inc(i_ind1);
  end;
Result:=NewChaine;
end;

Function  TraitePortionGformule(Chaine : string; TOBFTab : TOB) : string;
var TOBL : TOB ;
    i_ind1,ipos : integer;
    Fonct, ChainePol, StVariable : string;
begin
Result:='';
if Chaine='' then exit;
i_ind1 := 1; ChainePol:=''; StVariable:='';
while i_ind1 <= Length(Chaine) do
  begin
  if Chaine[i_ind1] = '[' then
    begin
    StVariable:='';
    ipos:=Pos(']',Copy(Chaine,i_ind1,Length(Chaine)));
    if ipos>0 then
      begin
      StVariable:=Copy(Chaine,i_ind1+1,ipos-2);
      StVariable:=AdapteVariable(StVariable);
      i_ind1:=i_ind1+ipos-1;
      end;
    end
  else if Chaine[i_ind1] = '@' then
    begin
    StVariable:='';
    ipos:=Pos('#',Copy(Chaine,i_ind1,Length(Chaine)));
    if ipos>0 then
      begin
      StVariable:=Copy(Chaine,i_ind1+1,ipos-2);
      i_ind1:=i_ind1+ipos-1;
      end;
    end
  else if Chaine[i_ind1] = '+' then fonct:='C60'
  else if Chaine[i_ind1] = '-' then fonct:='C70'
  else if Chaine[i_ind1] = '*' then fonct:='C50'
  else if Chaine[i_ind1] = '/' then fonct:='C10'
  else if Chaine[i_ind1] = ' ' then
  else if Chaine[i_ind1] <> '' then StVariable:=StVariable+Chaine[i_ind1];

  if fonct<>'' then
    begin
    TOBL:=TobFTab.FindFirst(['CO_CODE'],[Fonct],False);
    if TOBL<>Nil then
      begin
      if ChainePol='' then ChainePol:=TOBL.GetValue('CO_LIBELLE')+'('+StVariable
                      else if StVariable<>'' then ChainePol:=ChainePol+';'+StVariable;
      StVariable:='';
      fonct:='';
      end;
    end;
  Inc(i_ind1);
  end;
if ChainePol<>'' then
  if StVariable<>'' then ChainePol:=ChainePol+';'+StVariable+')'
                    else ChainePol:=ChainePol+')';
Result:=ChainePol;
end;

{==============================================================================================}
{=================================== Fonction d'appels ========================================}
{==============================================================================================}
Function  GFormuleVersFormatPolonais(Chaine : string) : string;
var TobFTab : TOB ;
    OuvPar, i_ind1 : integer;
    StArgument : string;
    QQ : TQuery;
begin
Result:='';
OuvPar := 0;
if Chaine='' then exit;
if Not VerifieParentheses(Chaine,True) then exit;
if Not VerifieCrochets(Chaine,True) then exit;
TobFTab:=Tob.Create('',Nil,-1);
QQ:=OpenSql('Select * from Commun Where CO_TYPE="GFX"',True);
if not QQ.Eof then TobFTab.LoadDetailDB('COMMUN','','',QQ,False);
Ferme(QQ);
i_ind1 := 1;
Chaine:=StringReplace(Chaine,',','.',[rfReplaceAll]);
Chaine:=ReStructureGformule(Chaine);
Chaine:=StringReplace(Chaine,'(','<',[rfReplaceAll]);
Chaine:=StringReplace(Chaine,')','>',[rfReplaceAll]);
while i_ind1 <= Length(Chaine) do
  begin
  if Chaine[i_ind1] = '<' then
    begin
    OuvPar:=i_ind1;
    end
  else if Chaine[i_ind1] = '>' then
    begin
    StArgument:=Copy(Chaine,OuvPar+1,i_ind1-OuvPar-1);
    StArgument:='@'+TraitePortionGformule(StArgument,TOBFTab)+'#';
    Chaine:=Copy(Chaine,1,OuvPar-1)+StArgument+Copy(Chaine,i_ind1+1,Length(Chaine));
    i_ind1:=0;
    end;
  Inc(i_ind1);
  end;
Result:=TraitePortionGformule(Chaine,TOBFTab);
TobFTab.Free;
end;

Function Saisie_FonctionParam(Formule : string) : R_FonctCal;
Var St : String;
begin
St:=AGLLanceFiche('GC','GCFONCTIONPARAM','','',Formule);
Result:=RecupVariableFormule(St);
end;

Function  RecupVariableFormule(Formule : string) : R_FonctCal;
begin
Result:=EvaluerFormule('',Formule,Nil,True);
end;

Function  EvaluationFormule(Titre,Formule : string) : R_FonctCal;
begin
Result:=EvaluerFormule(Titre,Formule,Nil,False);
end;

Function  EvaluationFormule(Titre : string; RFonct : R_FonctCal; ModeSilence : boolean=False) : R_FonctCal;
Var ind : integer;
    TOBVal,TOBL : TOB;
begin
TOBVal:=Tob.Create('Les Valeurs',Nil,-1);
TOBVal.AddChampSupValeur('AFFICHAGE',RFonct.Affichage);
for ind:=0 to High(RFonct.VarLibelle) do
  begin
  if RFonct.VarLibelle[ind]<>'' then
    begin
    TOBL:=Tob.Create('une Valeur',TOBVal,-1);
    TOBL.AddChampSupValeur('LIBELLE',RFonct.VarLibelle[ind]);
    TOBL.AddChampSupValeur('VALEUR',RFonct.VarValeur[ind]);
    if RFonct.Affichable[ind] then TOBL.AddChampSupValeur('AFFICHABLE','X')
                              else TOBL.AddChampSupValeur('AFFICHABLE','-');
    if RFonct.Modifiable[ind] then TOBL.AddChampSupValeur('MODIFIABLE','X')
                              else TOBL.AddChampSupValeur('MODIFIABLE','-');
    end;
  end;
Result:=EvaluerFormule(Titre,RFonct.Formule,TOBVal,ModeSilence);
TOBVal.Free;
end;

end.
