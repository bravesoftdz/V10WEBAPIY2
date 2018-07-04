unit AssistUtilBTP;

interface

uses    SysUtils, UTOB, HEnt1;

Function  Traite_Si(SiAlorsSinon, NomsTable, NomsFichier : string;
                    var Mess : string; var TSI : TOB) : boolean;
Function  Analyse_Condition(Condition, NomsTable, NomsFichier : string;
                            var Mess : string; var TSI : TOB) : boolean;
Function  Analyse_Vrai_Faux(VraiFaux, NomsTable, NomsFichier : string;
                            var Mess : string; var TSI : TOB) : boolean;
Function  Analyse_Condition_Simple(Condition, NomsTable, NomsFichier : string;
                                   var Mess : string; var TSI : TOB) : boolean;
Function  Analyse_Condition_Complexe(Condition, NomsTable, NomsFichier : string;
                                     var Mess : string; var TSI : TOB) : boolean;
Function  Analyse_Vrai_Faux_Valeur(VraiFaux, NomsTable, NomsFichier : string;
                                   var Mess : string; var TSI : TOB) : boolean;
Function  Analyse_Vrai_Faux_Valeur_Parenthese(Condition, NomsTable, NomsFichier : string;
                                              var Mess : string; var TSI : TOB) : boolean;
Function  Analyse_Vrai_Faux_Valeur_Operateur(Condition, NomsTable, NomsFichier : string;
                                             var Mess : string; var TSI : TOB) : boolean;
Function  VerifieElements(TCond : TOB; NomsTable, NomsFichier : string;
                          var Mess : string) : boolean;
procedure Eclate_Si(SiAlorsSinon : string; var Condition, Vrai, Faux : string);
Function  Verifie_Parentheses(Chaine : string; var Mess : string) : boolean;

const
	// libellés des messages
	Messages : array[0..6] of string = (
          {0}         'L''opérateur de comparaison est absent ou incorrect'
          {1}        ,'Le premier élément d''une condition est manquant'
          {2}        ,'Le deuxième élément d''une condition est manquant'
          {3}        ,'Le premier élément d''une condition n''est pas un champs'
          {4}        ,'Le deuxième élément d''une condition est incorrect'
          {5}        ,'Il y a une erreur dans les parenthèses'
          {6}        ,'La valeur de retour du ALORS ou du SINON est incorrecte'
                     );
        MotsCles : string = 'REPLACECAR;COMPTEUR;CORRESPOND;MAJUSCULE;TABLE;SOUSCHAINE;SPACE;STRINT;INTSTR;DATE';

implementation

Function  Traite_Si(SiAlorsSinon, NomsTable, NomsFichier : string;
                    var Mess : string; var TSI : TOB) : boolean;
var
    TobSI, TobALORS, TobSINON : TOB;
    Condition, Vrai, Faux : string;
    i_ind1, PosSi, PosAlors, PosSinon : integer;

begin
Result := True;
Eclate_Si(SiAlorsSinon, Condition, Vrai, Faux);
if TSI = nil then TSI := TOB.Create('', nil, -1);
TobSI := TOB.Create('', TSI, 0);
TobSI.AddChampSup('Type', True);
TobSI.AddChampSup('Valeur', True);
TobSI.PutValue('Type', 'Titre');
TobSI.PutValue('Valeur', 'SI');
if Vrai <> '' then
    begin
    TobALORS := TOB.Create('', TSI, 1);
    TobALORS.AddChampSup('Type', True);
    TobALORS.AddChampSup('Valeur', True);
    TobALORS.PutValue('Type', 'Titre');
    TobALORS.PutValue('Valeur', 'ALORS');
    end;
if Faux <> '' then
    begin
    TobSINON := TOB.Create('', TSI, 2);
    TobSINON.AddChampSup('Type', True);
    TobSINON.AddChampSup('Valeur', True);
    TobSINON.PutValue('Type', 'Titre');
    TobSINON.PutValue('Valeur', 'SINON');
    end;
Result := Analyse_Condition(Condition, NomsTable, NomsFichier, Mess, TobSI);
if not Result then Exit;
Result := Analyse_Vrai_Faux(Vrai, NomsTable, NomsFichier, Mess, TobALORS);
if not Result then Exit;
Result := Analyse_Vrai_Faux(Faux, NomsTable, NomsFichier, Mess, TobSINON);
end;

Function  Analyse_Condition(Condition, NomsTable, NomsFichier : string;
                            var Mess : string; var TSI : TOB) : boolean;
begin
if TSI = nil then TSI := TOB.Create('', nil, -1);
if Pos(' ET', Condition) + Pos(' OU', Condition) = 0 then
    Result := Analyse_Condition_Simple(Condition, NomsTable, NomsFichier, Mess, TSI)
    else
    Result := Analyse_Condition_Complexe(Condition, NomsTable, NomsFichier, Mess, TSI);
end;

Function  Analyse_Vrai_Faux(VraiFaux, NomsTable, NomsFichier : string;
                            var Mess : string; var TSI : TOB) : boolean;
var
    TobTemp, TobTemp2 : TOB;
    Condition, Vrai, Faux : string;

begin
Result := True;
if Pos('{SI', VraiFaux) <> 0 then
    begin
    TobTemp := TOB.Create('', TSI, -1);
    TobTemp.AddChampSup('Type', True);
    TobTemp.AddChampSup('Valeur', True);
    TobTemp.PutValue('Type', 'Titre');
    TobTemp.PutValue('Valeur', 'SI IMBRIQUE');
    Eclate_Si(VraiFaux, Condition, Vrai, Faux);
    TobTemp2 := TOB.Create('', TobTemp, 0);
    TobTemp2.AddChampSup('Type', True);
    TobTemp2.AddChampSup('Valeur', True);
    TobTemp2.PutValue('Type', 'Titre');
    TobTemp2.PutValue('Valeur', 'SI');
    Result := Analyse_Condition(Condition, NomsTable, NomsFichier, Mess, TobTemp2);
    if not Result then Exit;
    if Vrai <> '' then
        begin
        TobTemp2 := TOB.Create('', TobTemp, 1);
        TobTemp2.AddChampSup('Type', True);
        TobTemp2.AddChampSup('Valeur', True);
        TobTemp2.PutValue('Type', 'Titre');
        TobTemp2.PutValue('Valeur', 'ALORS');
        Result := Analyse_Vrai_Faux(Vrai, NomsTable, NomsFichier, Mess, TobTemp2);
        end;
    if not Result then Exit;
    if Faux <> '' then
        begin
        TobTemp2 := TOB.Create('', TobTemp, 2);
        TobTemp2.AddChampSup('Type', True);
        TobTemp2.AddChampSup('Valeur', True);
        TobTemp2.PutValue('Type', 'Titre');
        TobTemp2.PutValue('Valeur', 'SINON');
        Result := Analyse_Vrai_Faux(Faux, NomsTable, NomsFichier, Mess, TobTemp2);
        end;
    end
    else
    begin
    Result := Analyse_Vrai_Faux_Valeur(VraiFaux, NomsTable, NomsFichier, Mess, TSI);
    end;
end;

Function  Analyse_Condition_Simple(Condition, NomsTable, NomsFichier : string;
                                   var Mess : string; var TSI : TOB) : boolean;
var
    TobTemp, TobTemp2 : TOB;
    PosOper : integer;
    Ope, Elt1, Elt2 : string;

begin
Mess := '';
Result := True;
//  On elimine le cas de ET et de OU
if (Condition = ' ET ') or (Condition = ' OU ') then Exit;
//  Vérification de l'opérateur
//PosOper := Pos(' =', Condition) + Pos('<>', Condition) + Pos(' >', Condition) +
//           Pos('>=', Condition) + Pos(' <', Condition) + Pos('<=', Condition);
PosOper := Pos('<>', Condition);
if PosOper = 0 then
    begin
    PosOper := Pos('>=', Condition);
    if PosOper = 0 then
        Begin
        PosOper := Pos('<=', Condition);
        if PosOper = 0 then
            Begin
            PosOper := Pos(' =', Condition);
            if PosOper = 0 then
                Begin
                PosOper := Pos(' >', Condition);
                if PosOper = 0 then
                    PosOper := Pos(' <', Condition);
                end;
            end;
        end;
    end;
if PosOper = 0 then
    begin
    Mess := Messages[0];
    Result := False;
    Exit;
    end;
//  Récuperation des éléments de la condition
//PosOper := Pos(' = ', Condition) + Pos('<> ', Condition) + Pos('> ', Condition) +
//           Pos('>= ', Condition) + Pos('< ', Condition) + Pos('<= ', Condition);
Elt1 := Trim(Copy(Condition, 0, PosOper - 1));
Ope  := Trim(Copy(Condition, PosOper, 2));
Elt2 := Trim(Copy(Condition, PosOper + 2, Length(Condition) - 1));
//  Les éléments doivent être renseignés
if Elt1 = '' then
    begin
    Mess := Messages[1];
    Result := False;
    Exit;
    end;
if Elt2 = '' then
    begin
    Mess := Messages[2];
    Result := False;
    Exit;
    end;
//  Le premier élément doit être obligatoirement un nom de champ de la table ou du fichier
if Pos(Elt1, NomsTable) + Pos(Elt1, NomsFichier) + Pos(Copy(Elt2, 0, Pos('(', Elt2)-1), MotsCles) = 0 then
    begin
    Mess := Messages[3];
    Result := False;
    Exit;
    end;
//  le deuxième élément peut être un nom de champ de la table ou du fichier, les chaines
//  VRAI et FAUX, ou une valeur entre guillemets
if Pos(Elt2, NomsTable) = 0 then
    if Pos(Elt2, NomsFichier) = 0 then
        if Pos(Copy(Elt2, 0, Pos('(', Elt2)-1), MotsCles) = 0 then
            if Elt2 <> 'VRAI' then
                if Elt2 <> 'FAUX' then
                   if Elt2 <> 'VIDE' then
                    begin
                    if (Elt2[1] <> '"') and (Elt2[Length(Elt2)] <> '"') then
                        begin
                        Mess := Messages[4];
                        Result := False;
                        Exit;
                        end;
                    end;
//
TobTemp := TOB.Create('', TSI, -1);
TobTemp.AddChampSup('Type', True);
TobTemp.AddChampSup('Valeur', True);
TobTemp.PutValue('Type', 'Titre');
TobTemp.PutValue('Valeur', 'Condition');
TobTemp2 := TOB.Create('', TobTemp, 0);
TobTemp2.AddChampSup('Type', True);
TobTemp2.AddChampSup('Valeur', True);
TobTemp2.PutValue('Type', 'Valeur');
TobTemp2.PutValue('Valeur', Elt1);
TobTemp2 := TOB.Create('', TobTemp, 1);
TobTemp2.AddChampSup('Type', True);
TobTemp2.AddChampSup('Valeur', True);
TobTemp2.PutValue('Type', 'Operateur');
TobTemp2.PutValue('Valeur', Ope);
TobTemp2 := TOB.Create('', TobTemp, 2);
TobTemp2.AddChampSup('Type', True);
TobTemp2.AddChampSup('Valeur', True);
TobTemp2.PutValue('Type', 'Valeur');
TobTemp2.PutValue('Valeur', Elt2);
end;

Function  Analyse_Condition_Complexe(Condition, NomsTable, NomsFichier : string;
                                     var Mess : string; var TSI : TOB) : boolean;
var
    TobTemp : TOB;
    st1, st2, st3 : string;
    i_ind1, nbOuvPar, nbFerPar : integer;
    MotResa : boolean;

begin
Mess := '';
Result := True;
//  Controle des nombres de parentheses ouvertes et fermees
st1 := Trim(Condition);
Result := Verifie_Parentheses(st1, Mess);
if not Result then Exit;
//  Eclatement de la condition en sous_condition
TobTemp := TSI;
MotResa := False;
nbOuvPar := 0;
nbFerPar := 0;
i_ind1 := 1;
while i_ind1 <= Length(st1) do
    begin
//  Parenthese ouvrante : on cree un sous niveau de tob...dont on remontera à la
//  prochaine parenthese fermante. Ceci pour le 1er caractere de la chaine. quand le
//  caractere precedent n'est pas une parenthese ouvrante on remonte d'un cran pour
//  ouvrir une nouvelle fille de meme niveau
    if st1[i_ind1] = '(' then
        begin
        st3 := Copy(st2, Pos(' ',st2), 25);
        while Pos('<',st3) <> 0 do st3[Pos('<',st3)] := ' ';
        while Pos('>',st3) <> 0 do st3[Pos('>',st3)] := ' ';
        while Pos('=',st3) <> 0 do st3[Pos('=',st3)] := ' ';
        st3 := TrimLeft(st3);
        if Pos(st3, MotsCles) = 0 then
            begin
            Inc(nbOuvPar);
            if (i_ind1 > 1) then
                if (st1[i_ind1 - 1] <> '(') then
                begin
                if (Trim(st2) = 'ET') or (Trim(st2) = 'OU') then
                    begin
                    TobTemp.PutValue('Type', 'Logique');
                    TobTemp.PutValue('Valeur', st2);
                    end
                    else
                    Analyse_Condition_Simple(st2, NomsTable, NomsFichier, Mess, TobTemp);
                st2 := '';
                TobTemp := TobTemp.Parent;
                end;
            TobTemp := TOB.Create('', TobTemp, -1);
            TobTemp.AddChampSup('Type', True);
            TobTemp.AddChampSup('Valeur', True);
            TobTemp.PutValue('Type', 'Titre');
            TobTemp.PutValue('Valeur', '(');
            end
            else
            begin
            MotResa := True;
//  caractere normal : on le stocke dans la tob fille courante...
            st2 := st2 + st1[i_ind1];
            end;
        end
//  Parenthese fermante. on remonte d'un cran pour redescendre immediatement si le
//  caractere suivant est autre chose qu'une parenthese fermante
    else if st1[i_ind1] = ')' then
        begin
        if not MotResa then
            begin
            Inc(nbFerPar);
            if st1[i_ind1 - 1] <> ')' then
                begin
    //            TobTemp.PutValue('Type', 'Condition');
    //            TobTemp.PutValue('Valeur', st2);
                if (Trim(st2) = 'ET') or (Trim(st2) = 'OU') then
                    begin
                    TobTemp.PutValue('Type', 'Logique');
                    TobTemp.PutValue('Valeur', st2);
                    end
                    else
                    Analyse_Condition_Simple(st2, NomsTable, NomsFichier, Mess, TobTemp);
                st2 := '';
                end;
            TobTemp := TobTemp.Parent;
            if (i_ind1 < Length(st1)) then
                if (st1[i_ind1 + 1] <> ')') then
                begin
                TobTemp := TOB.Create('', TobTemp, -1);
                TobTemp.AddChampSup('Type', True);
                TobTemp.AddChampSup('Valeur', True);
                TobTemp.PutValue('Type', 'Titre');
                TobTemp.PutValue('Valeur', '(');
                end;
            end
            else
            begin
            MotResa := False;
//  caractere normal : on le stocke dans la tob fille courante...
            st2 := st2 + st1[i_ind1];
            st1[i_ind1] := ' ';
            end;
        end
    else
//  caractere normal : on le stocke dans la tob fille courante...
        st2 := st2 + st1[i_ind1];
    Inc(i_ind1);
    end;
//  Verification de chaque sous condition élémentaire
//Result := VerifieElements(TSI, NomsTable, NomsFichier, Mess);
end;

Function  Analyse_Vrai_Faux_Valeur(VraiFaux, NomsTable, NomsFichier : string;
                                   var Mess : string; var TSI : TOB) : boolean;
var
    TobTemp : TOB;
    st1 : string;
    nbOuvPar, nbFerPar, PosOper : integer;

begin
Result := True;
if VraiFaux = '' then Exit;
//  Verification des parentheses eventuelles
Result := Verifie_Parentheses(VraiFaux, Mess);
if not Result then Exit;
//  Si la fille contient un opérateur ou des parentheses, on l'éclate en sous filles
PosOper := Pos('*', VraiFaux) + Pos('/', VraiFaux) + Pos('-', VraiFaux) + Pos('+', VraiFaux);
if Pos('(', VraiFaux) <> 0 then
    begin
    Analyse_Vrai_Faux_Valeur_Parenthese(VraiFaux, NomsTable, NomsFichier, Mess, TSI);
    end
    else
    begin
    if PosOper <> 0 then
        begin
        Analyse_Vrai_Faux_Valeur_Operateur(VraiFaux, NomsTable, NomsFichier, Mess, TSI);
        end
        else
        begin
//  Si la fille ne contient pas un opérateur, on verifie qu'elle ne contient
//  que un champ de table, un champ du fichier, une constante entourée de guillemets ou un mot cle
//  Creation de la fille correspondant à la valeur
        TobTemp := TOB.Create('', TSI, -1);
        TobTemp.AddChampSup('Type', True);
        TobTemp.AddChampSup('Valeur', True);
        TobTemp.PutValue('Type', 'Valeur');
        TobTemp.PutValue('Valeur', VraiFaux);
        if Pos('(', VraiFaux) <> 0 then
            st1 := Copy(VraiFaux, 0, Pos('(', VraiFaux) - 1)
            else
            st1 := VraiFaux;
        if Pos(st1, MotsCles) <> 0 then Exit;
        if Pos(Trim(VraiFaux), NomsTable) <> 0 then Exit;
        if Pos(Trim(VraiFaux), NomsFichier) <> 0 then Exit;
        if (VraiFaux[1] = '"') and (VraiFaux[Length(VraiFaux)] = '"') then Exit;
        if IsNumeric(VraiFaux) then Exit;
        Result := False;
        Mess := Messages[6];
        TobTemp.Free;
        end;
    end;
end;

Function  Analyse_Vrai_Faux_Valeur_Parenthese(Condition, NomsTable, NomsFichier : string;
                                              var Mess : string; var TSI : TOB) : boolean;
var
    TobTemp : TOB;
    st1, st2, st3 : string;
    i_ind1, i_ind2, PosOper, nbOuvPar, nbFerPar : integer;

begin
Mess := '';
Result := True;
//  Controle des nombres de parentheses ouvertes et fermees
st1 := Trim(Condition);
//  Eclatement de la condition en sous_condition
TobTemp := TSI;
nbOuvPar := 0;
nbFerPar := 0;
i_ind1 := 1;
while i_ind1 <= Length(st1) do
    begin
//  Parenthese ouvrante : on cree un sous niveau de tob...dont on remontera à la
//  prochaine parenthese fermante. Ceci pour le 1er caractere de la chaine. quand le
//  caractere precedent n'est pas une parenthese ouvrante on remonte d'un cran pour
//  ouvrir une nouvelle fille de meme niveau
    if st1[i_ind1] = '(' then
        begin
//  gestion de l'utilisation de fcts predefinies dans le ALORS ou le SINON
        st3 := Copy(st1, i_ind1 - 8, 8);
        if Pos(st3, MotsCles) <> 0 then
            begin
            for i_ind2 := i_ind1 to 255 do
                begin
                st2 := st2 + st1[i_ind2];
                if st1[i_ind2] = ')' then Break;
                end;
                i_ind1 := i_ind2;
            end
            else
            begin
            Inc(nbOuvPar);
            if (i_ind1 > 1) then
                if (st1[i_ind1 - 1] <> '(') then
                begin
                PosOper := Pos('*', st2) + Pos('/', st2) + Pos('-', st2) + Pos('+', st2) + Pos('^', st2);
                if PosOper <> 0 then
                    begin
                    Analyse_Vrai_Faux_Valeur_Operateur(st2, NomsTable, NomsFichier, Mess, TobTemp);
                    end
                    else
                    begin
                    TobTemp.PutValue('Type', 'Valeur');
                    TobTemp.PutValue('Valeur', st2);
                    TobTemp := TobTemp.Parent;
                    end;
                end;
            st2 := '';
            TobTemp := TOB.Create('', TobTemp, -1);
            TobTemp.AddChampSup('Type', True);
            TobTemp.AddChampSup('Valeur', True);
            TobTemp.PutValue('Type', 'Titre');
            TobTemp.PutValue('Valeur', '(');
            end;
        end
//  Parenthese fermante. on remonte d'un cran pour redescendre immediatement si le
//  caractere suivant est autre chose qu'une parenthese fermante
    else if st1[i_ind1] = ')' then
        begin
        Inc(nbFerPar);
        if st1[i_ind1 - 1] <> ')' then
            begin
            PosOper := Pos('*', st2) + Pos('/', st2) + Pos('-', st2) + Pos('+', st2) + Pos('^', st2);
            if PosOper <> 0 then
                begin
                Analyse_Vrai_Faux_Valeur_Operateur(st2, NomsTable, NomsFichier, Mess, TobTemp);
                end
                else
                begin
                TobTemp.PutValue('Type', 'Valeur');
                TobTemp.PutValue('Valeur', st2);
                end;
            st2 := '';
            end;
        TobTemp := TobTemp.Parent;
        if (i_ind1 < Length(st1)) then
            if (st1[i_ind1 + 1] <> ')') then
            begin
            TobTemp := TOB.Create('', TobTemp, -1);
            TobTemp.AddChampSup('Type', True);
            TobTemp.AddChampSup('Valeur', True);
            if (Pos(st1[i_ind1 + 1], '+-*/^') <> 0) or (Pos(st1[i_ind1 + 2], '+-*/^') <> 0) then
                begin
                if Pos(st1[i_ind1 + 1], '+-*/^') <> 0 then
                    begin
                    st2 := st1[i_ind1 + 1];
                    Inc(i_ind1);
                    end
                    else if Pos(st1[i_ind1 + 2], '+-*/^') <> 0 then
                        begin
                        st2 := st1[i_ind1 + 2];
                        i_ind1 := i_ind1 + 2;
                        end;
                TobTemp.PutValue('Type', 'Operateur');
                TobTemp.PutValue('Valeur', st2);
                TobTemp := TobTemp.Parent;
//                TobTemp := TOB.Create('', TobTemp, -1);
//                TobTemp.AddChampSup('Type', True);
//                TobTemp.AddChampSup('Valeur', True);
                st2 := '';
                end;
            end;
        end
    else
//  caractere normal : on le stocke dans la tob fille courante...
        st2 := st2 + st1[i_ind1];
    Inc(i_ind1);
    end;
if st2 <> '' then
    begin
    PosOper := Pos('*', st2) + Pos('/', st2) + Pos('-', st2) + Pos('+', st2) + Pos('^', st2);
    if PosOper <> 0 then
        Analyse_Vrai_Faux_Valeur_Operateur(st2, NomsTable, NomsFichier, Mess, TobTemp)
        else
        begin
        if TobTemp = TSI then
            begin
            TobTemp := TOB.Create('', TobTemp, -1);
            TobTemp.AddChampSup('Type', True);
            TobTemp.AddChampSup('Valeur', True);
            end;
        TobTemp.PutValue('Type', 'Valeur');
        TobTemp.PutValue('Valeur', st2);
        end;
    end;
end;

Function  Analyse_Vrai_Faux_Valeur_Operateur(Condition, NomsTable, NomsFichier : string;
                                             var Mess : string; var TSI : TOB) : boolean;
var
    TobTemp : TOB;
    st1, st2 : string;
    i_ind1 : integer;

begin
Mess := '';
Result := True;
//  Controle des nombres de parentheses ouvertes et fermees
st1 := Trim(Condition);
//  Eclatement de la condition en sous_condition
TobTemp := TSI;
i_ind1 := 1;
while i_ind1 <= Length(st1) do
    begin
//  Premier caractere : on cree un sous niveau de tob...dont on remontera au
//  prochain opérateur rencontré
    if i_ind1 = 1 then
        begin
        TobTemp := TOB.Create('', TobTemp, -1);
        TobTemp.AddChampSup('Type', True);
        TobTemp.AddChampSup('Valeur', True);
        end;
//  Operateur : on enregistre l'opérande précedent dans sa tob, on remonte d'un cran pour
//  ouvrir une nouvelle fille de meme niveau pour l'opérateur puis on recree une autre
//  fille pour l'operande suivant
    if ((Pos(st1[i_ind1], '+-*/^') <> 0) and ((st1[i_ind1 - 1] <> '"') or (st1[i_ind1 + 1] <> '"'))) and
        (st1[i_ind1 - 1] <> st1[i_ind1 + 1]) then
        begin
        if i_ind1 > 1 then
            begin
            TobTemp.PutValue('Type', 'Valeur');
            TobTemp.PutValue('Valeur', st2);
            st2 := '';
            TobTemp := TobTemp.Parent;
            end;
        TobTemp := TOB.Create('', TobTemp, -1);
        TobTemp.AddChampSup('Type', True);
        TobTemp.AddChampSup('Valeur', True);
        TobTemp.PutValue('Type', 'Operateur');
        TobTemp.PutValue('Valeur', st1[i_ind1]);
        TobTemp := TobTemp.Parent;
        TobTemp := TOB.Create('', TobTemp, -1);
        TobTemp.AddChampSup('Type', True);
        TobTemp.AddChampSup('Valeur', True);
        end
    else
//  caractere normal : on le stocke dans la tob fille courante...
        st2 := st2 + st1[i_ind1];
    Inc(i_ind1);
    end;
TobTemp.PutValue('Type', 'Valeur');
TobTemp.PutValue('Valeur', st2);
end;

Function  VerifieElements(TCond : TOB; NomsTable, NomsFichier : string;
                          var Mess : string) : boolean;
var
    TobTemp : TOB;
    i_ind1 : integer;
    st1 : string;

begin
Result := True;
for i_ind1 := 0 to TCond.Detail.Count - 1 do
    begin
    TobTemp := TCond.Detail[i_ind1];
    st1 := Trim(TobTemp.GetValue('Valeur'));
    if TobTemp.Detail.Count > 0 then
        Result := VerifieElements(TobTemp, NomsTable, NomsFichier, Mess)
        else
        Result := Analyse_Condition_Simple(st1, NomsTable, NomsFichier, Mess, TobTemp);
    end;
    if not Result then Exit;
end;

procedure Eclate_Si(SiAlorsSinon : string; var Condition, Vrai, Faux : string);
var
    PosSi, PosAlors, PosSinon,Pos1Croch : integer;
    i_ind1, i_ind2, nbOuvCroch, nbFerCroch : integer;
    st1, st2 : string;
    NewSi : boolean;

begin
//  Verification de l'existence et de la validite des crochets delimiteurs de SI
nbOuvCroch := 0;
nbFerCroch := 0;
i_ind1 := 1;
while i_ind1 <= Length(SiAlorsSinon) do
    begin
    if SiAlorsSinon[i_ind1] = '{' then
        Inc(nbOuvCroch)
    else if SiAlorsSinon[i_ind1] = '}' then
        Inc(nbFerCroch);
    Inc(i_ind1);
    end;
if (nbOuvCroch > 0) and (nbOuvCroch <> nbFerCroch) then
    begin
    Condition := '';
    Vrai := '';
    Faux := '';
    Exit;
    end;
//  on balaye le SI d'origine pour extraire le bloc ALORS et le bloc SINON
PosSi    := Pos('{SI'  , SiAlorsSinon);
PosAlors := Pos('ALORS', SiAlorsSinon);
PosSinon := Pos('SINON', SiAlorsSinon);
Pos1Croch := Pos('}', SiAlorsSinon);
if (nbOuvCroch = nbFerCroch) and (nbOuvCroch = 1) then
    begin
//  il y a 1 jeu de crochet, donc pas de SI imbriqués. on travaille à partir de
//  la position des chaines [SI, ALORS et SINON
    if PosSinon <> 0 then
        Faux := Trim(Copy(SiAlorsSinon, PosSinon + 5, (Length(SiAlorsSinon) - (PosSinon + 5))));
    if PosAlors <> 0 then
        if PosSinon = 0 then
            Vrai := Trim(Copy(SiAlorsSinon, PosAlors + 5, (Length(SiAlorsSinon) - (PosAlors + 5))))
            else
            Vrai := Trim(Copy(SiAlorsSinon, PosAlors + 5, (PosSinon - 1) - (PosAlors + 5)));
    end
    else
    begin
//  Il y a plusieurs jeu de crochets, donc des SI imbriqués. On recupere la condition
//  à partir des position du [SI et du ALORS. on recupere le bloc ALORS en balayant
//  à partir de la position de la chaine ALORS entre le [ et le ] de meme niveau,
//  sans se preoccuper des crochets intermediaires.
    if Pos1Croch < PosSinon then
        begin
        Vrai := Trim(Copy(SiAlorsSinon, PosAlors + 5, (PosSinon - 1) - (PosAlors + 5)));
        Faux := Trim(Copy(SiAlorsSinon, PosSinon + 5, (Length(SiAlorsSinon)) - (PosSinon + 5)));
        end
        else
        begin
        i_ind1 := PosAlors + 5;
        NewSi := False;
        while i_ind1 <= Length(SiAlorsSinon) do
            begin
            if SiAlorsSinon[i_ind1] = '{' then
                begin
                Inc(nbOuvCroch);
                NewSi := True;
                end
            else if SiAlorsSinon[i_ind1] = '}' then
                begin
                Inc(nbFerCroch);
                if nbOuvCroch = nbFerCroch then
                    begin
                    Vrai := st2 + SiAlorsSinon[i_ind1];
                    NewSi := False;
                    Break;
                    end;
                end;
            if NewSi then st2 := st2 + SiAlorsSinon[i_ind1];
            Inc(i_ind1);
            end;
    //  On a traité le bloc ALORS, on verifie maintenant qu'il existe un SINON dans les
    //  10 caractères suivant. Si oui, on extrait le bloc SINON
            st2 := Trim(Copy(SiAlorsSinon, i_ind1 + 1, 10));
            if Pos('SINON', st2) <> 0 then
                begin
                st2 := Trim(Copy(SiAlorsSinon, i_ind1 + 1, (Length(SiAlorsSinon) - (i_ind1 + 1))));
                Faux := Trim(Copy(st2, Pos('SINON', st2) + 5, Length(st2)));
                end;
        end;
    end;
Condition := Trim(Copy(SiAlorsSinon, PosSi + 3, (PosAlors - 1) - (PosSi + 3)));
end;

Function  Verifie_Parentheses(Chaine : string; var Mess : string) : boolean;
var
    nbOuvPar, nbFerPar, i_ind1 : integer;

begin
Result := True;
nbOuvPar := 0;
nbFerPar := 0;
i_ind1 := 1;
while i_ind1 <= Length(Chaine) do
    begin
    if Chaine[i_ind1] = '(' then
        Inc(nbOuvPar)
    else if Chaine[i_ind1] = ')' then
        Inc(nbFerPar);
    Inc(i_ind1);
    end;
if nbOuvPar <> nbFerPar then
    begin
    Mess := Messages[5];
    Result := False;
    Exit;
    end;
end;

end.
