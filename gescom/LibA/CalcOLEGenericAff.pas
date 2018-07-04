{***********UNITE*************************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 27/03/2003
Modifié le ... :   /  /    
Description .. : Source qui comporte les fonctions utilisables dna sles états 
Suite ........ : de la GI/GA et qui sont génériques
Suite ........ : Ont  besoin d'info des paramsoc ou fonctionne seule
Suite ........ : Ce source est fait à partir de fct qui existaient dans d'autres 
Suite ........ : sources, et qui ont été mis dans celui_ci pour éviter des 
Suite ........ : include en cascade
Mots clefs ... : FONCTION;ETAT
*****************************************************************}
unit CalcOLEGenericAff;

interface
uses HEnt1, SysUtils, HmsgBox ,hctrls ,ParamSoc ;

function AFGenCalcOLEEtat(sf,sp : string) : variant ;
  // fct utilisées dans d'autres sources GI/GA.. 
Function CodeAffaireAffiche (CodeAffaire : String; Separateur: string='') : String;forward;
function AfReadParam(var St : String) : string ;
Function HeureBase100To60 ( Heure : Double) : Double;
Procedure CodeAffaireDecoupe(CodeAff : string; Var Part0,Part1, Part2, Part3, Avenant : String; FTypeAction:TActionFiche; SaisieAffaire:Boolean);
Function TestStatutReduit (stStatut : String) : String;
Function TestCodeAvenant (stAvenant : String; Alim00 : Boolean) : String;
Function OLEFormatPeriode (Periode : string; bLib : Boolean) :string;  

implementation
Function AfficheCodeAff (Part0,Part1,PArt2,Part3,Avenant : String ) : String; forward;
function OLEFormatExercice (Exer : string) : String; forward;
function AFLibExercice ( Exercice : string) : String; forward;


Function HeureBase100To60 ( Heure : Double) : Double;
var tmp : double;
BEGIN
if Heure = 0 then BEGIN Result := 0; Exit; END;
tmp := (Frac(Heure)*0.6);
if (tmp = 0) then Result := Trunc(Heure) else Result :=Trunc(Heure) + Arrondi(tmp,2);
END;

{***********Gestion des statuts************}
Function TestStatutReduit (stStatut : String) : String;
BEGIN
Result := stStatut;
if (stStatut <> 'A') and (stStatut <> 'P') then Result := 'A';
END;

{ *********Gestion des avenants + Etats ... ***********}
Function TestCodeAvenant (stAvenant : String; Alim00 : Boolean) : String;
BEGIN
 // mcd 28/03/03 change VH_GC en Paramsoc
Result := stAvenant;
if (GetParamSoc('SO_AFFGestionAvenant')) and (StAvenant = '') and Not(Alim00) then exit;
if (stAvenant < '01') or (stAvenant > '99') then Result := '00';
END;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/02/2000
Modifié le ... : 17/02/2000
Description .. : Décomposition du code affaire concaténé
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
Procedure CodeAffaireDecoupe(CodeAff : string; Var Part0, Part1, Part2, Part3, Avenant : String; FTypeAction:TActionFiche; SaisieAffaire:Boolean);
var i,lng,posit: integer;
    TypeCode : string3;
    Valeur,TexteSel,CodePartie : String;
Begin
    // mcd 28/03/03 passage de tous les VH_GC en GetPAramSoc
  Part1 := ''; Part2 := ''; Part3 := '';
  Posit := 0;  lng := 0;
  If (CodeAff = '') then Exit;
    // mcd 28/03/03 for i:=1 to NbMaxPartieAffaire do
  for i:=1 to 3 do
    Begin
    // recup des valeurs de la partie du code traitée
    case i of
        1: Begin lng:=GetParamSoc('SO_AFFCo1Lng');
                 TypeCode:=GetParamSoc('SO_AFFCo1Type');
                 Valeur:=GetParamSoc('SO_AFFCo1valeur');
                 posit:=2; End;
        // Attention le premier car. réservé pour le statut
        2: Begin lng:= GetParamSoc('SO_AFFCo2Lng');
                 TypeCode:=GetParamSoc('SO_AFFCo2Type');
                 GetParamSoc('SO_AFFCo2valeur');
                 posit:=GetParamSoc('SO_AFFCo1Lng')+2; End;
        3: Begin lng:=GetParamSoc('SO_AFFCo3Lng');
                 TypeCode:=GetParamSoc('SO_AFFCo3Type');
                 Valeur:=GetParamSoc('SO_AFFCo3valeur');
                 posit:=GetParamSoc('SO_AFFCo1Lng')+GetParamSoc('SO_AFFCo2Lng')+2;
                 End;
        End;
    // Attention  longueur Maxi : 15 (car 16 et 17 = avenant )
    if Posit + Lng > 16 then Lng := 16 -Posit;
    If (i <= GetParamSoc('SO_AFFCodeNbPartie')) Then
        Begin   // Partie utilisée
        TexteSel:='';
        If( CodeAff<> '') Then TexteSel:= trim(Copy (CodeAff,posit,lng));
        If (SaisieAffaire) then
            Begin
            If (FTypeAction=taCreat) Then CodePartie:= '' Else CodePartie:=TexteSel;
            End
        Else // Code affaire depuis une autre saisie
            Begin
            If (CodeAff<>'') Then CodePartie:=TexteSel Else CodePartie:= '';
            End;

        // spécificitées en fonction du type de zone ...
           // mcd 18/09/01 on ne passe dans cette fct que qd le code aff est renseigné .. il ne faut pas écraser ce qui ets mis dans la zone If (Typecode= 'FIX') Then   CodePartie:=Valeur Else          // Zone fixe
        If ((Typecode='CPT') And SaisieAffaire) Then                    // compteur
            Begin
            If (FTypeAction=taCreat) Then CodePartie:=Valeur;
            End;
        End
    Else
        Begin  // RAZ de la partie de structure non utilisée
        CodePartie:='';
        End;
    case i of
        1: Part1 := CodePartie;
        2: Part2 := CodePartie;
        3: Part3 := CodePartie;
        End;
    End;
  // Traitement du statut et de la zone avenant
  Part0 := TestStatutreduit(Copy(CodeAff,1,1));
  Avenant := TestCodeAvenant(Copy (codeAff,16,2),False);
End;


{***********A.G.L.***********************************************
Auteur  ...... : DESSEIGNET MC
Créé le ...... : 08/08/2000
Modifié le ... : 08/08/2000
Description .. : A partir des 5 partie de l'affaire,Retourne le code affaire
    à afficher / format et zones visibles
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
Function AfficheCodeAff (Part0,Part1,PArt2,Part3,Avenant : String ) : String;
Var Separateur: String;
BEGIN
Result := '';
 // mcd 18/09/01 même en GA on met un séparateur if (ctxScot in V_PGI.PGIContexte) then separateur := ' ';
  // mcd 28/03/03 passe tous les VH_GC en GetparamSoc
separateur := ' ';
Result := Part1;
if GetParamSoc('SO_AFFCo2Visible') then
   BEGIN
   if (GetParamSoc('SO_AFFORMATEXER')<>'AUC') then Part2 := AFLibExercice (Part2);
   Result := Result + separateur + Part2;
   END;
if GetParamSoc('SO_AFFCo3Visible') then Result := Result + separateur + Part3;
if (GetParamSoc('SO_AFFGestionAvenant')) And (Avenant <> '00') then
    BEGIN
    Result := Result + ' ' + Avenant;
    END;
END;

{***********A.G.L.***********************************************
Auteur  ...... : PATRICE ARANEGA
Créé le ...... : 21/05/2000
Modifié le ... : 21/05/2000
Description .. : Retourne le code affaire affiché / format et zones visibles
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
Function CodeAffaireAffiche (CodeAffaire : String; Separateur: string='') : String;
Var Part0,Part1, Part2, Part3,Avenant: String;
BEGIN
Result := '';
 // mcd 18/09/01 même en GA on met un séparateur if (ctxScot in V_PGI.PGIContexte) then separateur := ' ';
 // mcd 28/03/03 passage de tous les VH_GC en GetParamSoc
separateur := ' ';
CodeAffaireDecoupe(CodeAffaire, part0, Part1, Part2, Part3,Avenant,taConsult,False);
Result := Part1;
if GetParamSoc('SO_AFFCo2Visible') then
   BEGIN
   if (GetParamSoc('SO_AFFORMATEXER')<>'AUC') then Part2 := AFLibExercice (Part2);
   Result := Result + separateur + Part2;
   END;
if GetParamSoc('SO_AFFCo3Visible') then Result := Result + separateur + Part3;
if (GetParamSOc('SO_AFFGestionAvenant')) And (Avenant <> '00') then
    BEGIN
    Result := Result + ' ' + Avenant;
    END;
END;


function AFLibExercice ( Exercice : string) : String;
Var TypeExer : String;
BEGIN
  // mcd 28/03/03 passage de tous les VH_GC en GetPAramSoc
Result := Exercice;
TypeExer := GetParamSoc('SO_AFFormatExer');
if TypeExer = 'AUC' then Exit else
 if TypeExer = 'A' then Exit else
  if TypeExer = 'AM'  then
    BEGIN
    if Length(Exercice) < 6 then Exit;
    Result := Copy(Exercice,1,4) +'-'+ Copy(Exercice,5,2);
    END else
   if TypeExer = 'AA' then
     BEGIN
     if Length(Exercice) < 8 then Exit;
     Result := Copy(Exercice,1,4) +'-'+ Copy(Exercice,5,4);
     END;
END;


function OLEFormatExercice (Exer : string) : String;
BEGIN
Result := '';
       // mcd 18/09/01 if not (ctxScot in V_PGI.PGIContexte) then Exit;
if GetParamSoc('SO_AfFormatExer')='AUC' then Exit;
Trim(Exer); Result := Exer;
if (Length(Exer) < 5) then Exit;
Result := Copy(Exer,1,4)+'_'+ Copy(Exer, 5, Length(Exer)-4);
END;


Function OLEFormatPeriode (Periode : string; bLib : Boolean) :string;
Var stYear, StMois : string;
    IMois  : integer;
BEGIN
Result := '';
if (Periode = '') or (periode < '111111')  then exit;
stYear := Copy (Periode,1,4);
if bLib then
    BEGIN
    IMois  := StrToInt(Copy(Periode,5,2));
    Result := FirstMajuscule(ShortMonthNames[IMois])+' '+stYear;
    END
else
    BEGIN
    stMois := Copy (Periode,5,2);
    Result := stMois +'/'+ stYear;
    END;
END;

function AfReadParam(var St : String) : string ;
var i : Integer ;
begin
i:=Pos(';',St) ; if i<=0 then i:=Length(St)+1 ; result:=Copy(St,1,i-1) ; Delete(St,1,i) ;
Result := Trim(Result);
end ;

function AFGenCalcOLEEtat(sf,sp : string) : variant ;
Var st1, St2, St3, St4, St5  : string;
 bb : boolean;
 hr1 : TDateTime;
BEGIN
TVarData(Result).VType := varEmpty;
if sf='FORMATEXERCICE' then
   BEGIN
   st1 := AfReadParam(sp) ;
   if not (ctxScot in V_PGI.PGIContexte) then BEGIN Result := ' '; Exit; END;
   Result := OLEFormatExercice (st1);
   Exit;
   END else
if sf='FORMATCODEAFF' then
   BEGIN
   st1 := AfReadParam(sp) ;
   Result := CodeAffaireAffiche(st1);
   Exit;
   END else
if sf='AFFICHECODEAFF' then
   BEGIN
   st1  := AfReadParam(sp) ;
   st2  := AfReadParam(sp) ;
   st3  := AfReadParam(sp) ;
   st4  := AfReadParam(sp) ;
   st5  := AfReadParam(sp) ;
   Result := AfficheCodeAff(st1,st2,st3,st4,st5);
   Exit;
   END else
if sf='VALOPERMISE' then
    BEGIN
    If ExJaiLeDroitConcept ( cc32 , false) then Result := sp else Result := 0.0;
    END else
if sf='FORMATPERIODE' then
    BEGIN
    st1  := AfReadParam(sp) ;    // periode concaténée
    st2  := AfReadParam(sp) ;    // libellé ou mois numérique
    bb := (st2 = 'LIB');
    Result := OLEFormatPeriode (st1, bb);
    END else
//NCX 31/07/01
if sf='RECUPHEURE' then
    BEGIN
     hr1:= StrToFloat(AfReadParam(sp)) ;    // double à convertir
     st1 := FormatFloat('0.00',HeureBase100To60(hr1));
     st2 := ReadTokenPipe(st1,',');
     If ((st2='00') or (st2='0')) and (st1='00') then
     result := '' else result := st2+'h'+st1;
   END;

END;


end.
