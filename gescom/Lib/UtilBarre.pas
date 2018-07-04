unit UtilBarre;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, StdCtrls;

procedure ControleCodeBarre (CodeBarre : string; TypeCodeBarre : string;
                             var ValeurPoss : string; TypeControle : integer);

var i_somme : integer;

implementation

procedure Controle39C (CodeBarre : string; var ValeurPoss : string;
                       TypeControle : integer);
var i_ind : integer;
    i_ind2 : integer;
    i_reste : integer;
const IndiceCara : array [0..6] of char = ('-', '.', ' ', '$', '/', '+', '%');
Const NB_CARA_SPEC : integer = 6;
begin
    i_somme := 0;
    i_ind := length (CodeBarre);
    if i_ind <= 44  then
    begin
        if CodeBarre [i_ind] = '*' then i_ind := i_ind - 1;
    end else
    begin
        i_ind := -1;
    end;
    while (i_ind > 0) and (CodeBarre [i_ind] <> '*') do
    begin
        i_ind2 := NB_CARA_SPEC;
        while (i_ind2 >= 0) and  (IndiceCara [i_ind2] <> CodeBarre [i_ind]) do
        begin
            i_ind2 := i_ind2 - 1;
        end;
        if i_ind2 >= 0 then
        begin
            i_somme := i_somme + i_ind2 + 36;
        end
        else if ((CodeBarre [i_ind] >= '0') and (CodeBarre [i_ind] <= '9')) or
                ((CodeBarre [i_ind] >= 'A') and (CodeBarre [i_ind] <= 'Z')) or
                ((CodeBarre [i_ind] >= 'a') and (CodeBarre [i_ind] <= 'z')) then
        begin
            CodeBarre [i_ind] := Upcase (CodeBarre [i_ind]);
            i_somme := i_somme + (integer(CodeBarre [i_ind])) - 48;
            if (CodeBarre [i_ind] >= 'A') and (CodeBarre [i_ind] <= 'Z') then
            begin
                i_somme := i_somme - 7;
            end;
        end else
        begin
            i_ind := -1;
        end;
        i_ind := i_ind - 1;
    end;
    if (i_ind = 0) or ((i_ind = 1) and (CodeBarre [i_ind] = '*')) then
    begin
        i_reste := i_somme mod 43;
        if i_reste < 36 then
        begin
            if i_reste > 9 then
            begin
                if TypeControle = 0 then // proposition de type de code
                begin
                    ValeurPoss := 'Code 39' + chr (i_reste + 55);
                end else if TypeControle = 1 then
                begin // proposition de code pour le type choisi avec ajout d'un chiffre
                    ValeurPoss := '';
                end else if TypeControle = 2 then
                begin // controle uniquement entre type et code
                    ValeurPoss := '39';
                end;
            end else
            begin
                if TypeControle = 0 then
                begin
                    ValeurPoss := 'Code 39' + chr (i_reste + 48);
                end else if TypeControle = 1 then
                begin // proposition de code pour le type choisi
                    ValeurPoss := '';
                end else
                begin // controle uniquement entre type et code
                    ValeurPoss := '39';
                end;
            end;
        end else
        begin
            if TypeControle = 0 then
            begin
                ValeurPoss := 'Code 39' + IndiceCara [i_reste - 36];
            end else if TypeControle = 1 then
            begin // proposition de code pour le type choisi
                ValeurPoss := '';
            end else
            begin // controle uniquement entre type et code
                ValeurPoss := '39';
            end;
        end;
    end else
    begin
        i_ind := -1;
    end;
    if i_ind = -1 then
    begin
        ValeurPoss := '';
    end;
end;

procedure Controle39 (CodeBarre : string);
begin
end;

procedure Controle128A (CodeBarre : string; var ValeurPoss : string;
                        TypeControle : integer);
var i_ind : integer;
    i_reste : integer;
begin
    i_somme := 0;
    i_ind := length (CodeBarre) - 3;
    while i_ind > 1 do
    begin
        if (CodeBarre [i_ind] >= 'a') and (CodeBarre [i_ind] <= 'z') then
        begin
            CodeBarre [i_ind] := UpCase (CodeBarre [i_ind]);
        end;
        if (CodeBarre [i_ind] >= '!') and (CodeBarre [i_ind] <= 'Z') then
        begin
            i_somme := i_somme + (((integer(CodeBarre [i_ind])) - 32) * i_ind);
        end;
        i_ind := i_ind - 1;
    end;
    i_reste := i_somme mod 103;
    if (i_reste > 0) and (i_reste <= 63) then
    begin
        if TypeControle = 0 then
        begin
            ValeurPoss := '(Start A)Code 128(' + Chr (i_reste + 32)+')';
        end else if TypeControle = 1 then
        begin
            ValeurPoss := '';
        end else
        begin
            ValeurPoss := '';
        end;
    end else if i_reste = 0 then
    begin
        if TypeControle = 0 then
        begin
            ValeurPoss := '(Start A)Code 128';
        end else if TypeControle = 1 then
        begin
            ValeurPoss := '';
        end else
        begin
            ValeurPoss := '';
        end;
    end;
end;

procedure Controle128B (CodeBarre : string; var ValeurPoss : string;
                        TypeControle : integer);
var i_ind : integer;
    i_reste : integer;
begin
    i_somme := 0;
    i_ind := length (CodeBarre) - 3;
    while i_ind > 3 do
    begin
        if ((CodeBarre [i_ind] >= '!') and (CodeBarre [i_ind] <= '{')) or
           (CodeBarre [i_ind] = '}') then
        begin
            i_somme := i_somme + (((integer(CodeBarre [i_ind])) - 32) * i_ind);
        end;
        i_ind := i_ind - 1;
    end;
    i_reste := i_somme mod 103;
    if ((i_reste > 0) and (i_reste <= 91)) or (i_reste = 93) then
    begin
        if TypeControle = 0 then
        begin
            ValeurPoss := '(Start B)Code 128(' + Chr (i_reste + 32)+')';
        end else if TypeControle = 1 then
        begin
            ValeurPoss := '';
        end else
        begin
            ValeurPoss := '';
        end;
    end;
end;

procedure Controle128C (CodeBarre : string; var ValeurPoss : string;
                        TypeControle : integer);
begin
end;

procedure Controle128 (CodeBarre : string; var ValeurPoss : string;
                        TypeControle : integer);
var i_ind : integer;
begin
    i_ind := length (CodeBarre);
    if (CodeBarre [i_ind] = '6') and (CodeBarre [i_ind - 1] = '0') and
       (CodeBarre [i_ind - 2] = '1') then
    begin
        if (CodeBarre [1] = 'A') or (CodeBarre [1] = 'a') then
        begin
            Controle128A (CodeBarre, ValeurPoss, TypeControle);
        end else if (CodeBarre [1] = 'B') or (CodeBarre [1] = 'b') then
        begin
            Controle128B (CodeBarre, ValeurPoss, TypeControle);
        end else if (CodeBarre [1] = 'C') or (CodeBarre [1] = 'c') then
        begin
            Controle128C (CodeBarre, ValeurPoss, TypeControle);
        end else
        begin
            i_ind := -1;
        end;
    end else
    begin
        i_ind := -1;
    end;
    if i_ind = -1 then
    begin
        ValeurPoss := '';
    end;
end;                                     

procedure ControleEAN (CodeBarre : string; var ValeurPoss : string;
                       TypeControle : integer; TypeCode : integer);
// controle les codes barres de type E13
var i_ind : integer;
    i_longueur_code : integer;
    i_somme_odd : integer;
    i_somme_even : integer;
begin
    i_somme := 0;
    i_somme_odd := 0;
    i_somme_even := 0;
    i_longueur_code := length (CodeBarre);
    if (i_longueur_code < TypeCode - 1) or (i_longueur_code > TypeCode) then i_ind := -1
    else i_ind := TypeCode - 1;
    while i_ind > 0 do
    begin
        if ((CodeBarre [i_ind] < '0') or (CodeBarre [i_ind] > '9')) then
        begin
            i_ind := -1;
        end else
        begin
            if ((i_ind mod 2) = 0) and (i_ind <> 1) then
            begin
                if TypeCode = 13 then
                    i_somme_odd := i_somme_odd + strtoint (CodeBarre [i_ind])
                else i_somme_even := i_somme_even + strtoint (CodeBarre [i_ind]);
            end else
            begin
                if TypeCode = 13 then
                    i_somme_even := i_somme_even + strtoint (CodeBarre [i_ind])
                else i_somme_odd := i_somme_odd + strtoint (CodeBarre [i_ind]);
            end;
            i_ind := i_ind - 1;
        end;
    end;
    if i_ind <> -1 then
    begin
        if (i_longueur_code = TypeCode) and
           ((CodeBarre [TypeCode] < '0') or (CodeBarre [i_ind] > '9')) then
        begin
            i_ind := -1;
        end;
    end;
    if i_ind <> -1 then
    begin
        i_somme := (i_somme_odd * 3) + i_somme_even;
        i_ind := 0;
        while ((i_somme + i_ind) mod 10) <> 0 do i_ind := i_ind + 1;
        if TypeControle < 2 then
        begin
            if i_longueur_code < TypeCode then
            begin
                ValeurPoss := CodeBarre + chr (i_ind + 48);
            end else if i_ind <> strtoint (CodeBarre [TypeCode]) then
            begin
                ValeurPoss := CodeBarre;
                ValeurPoss [Typecode] := chr (i_ind + 48);
            end;
        end else if TypeControle = 2 then
        begin
            if (i_longueur_code < TypeCode) or
               (strtoint (CodeBarre [TypeCode]) <> i_ind) then
            begin
                ValeurPoss := '';
            end else
            begin
                if TypeCode < 10 then
                    ValeurPoss := 'E' + inttostr (TypeCode)
                else ValeurPoss := 'EA' + inttostr (TypeCode);
            end;
        end;
    end else
    begin
            ValeurPoss := '';
    end;
end;

procedure ControleITC (CodeBarre : string);
// controle les codes barres de type ITC
begin
    Application.MessageBox ('Code ITC', '', MB_OKCANCEL);
end;

procedure ControleITF (CodeBarre : string);
// controle les codes barres de type ITF
begin
    Application.MessageBox ('Code ITF', '', MB_OKCANCEL);
end;

procedure ControleCodeBarre (CodeBarre : string;
                             TypeCodeBarre : string;
                             var ValeurPoss : string;
                             TypeControle : integer);
// procedure de controle d'un code barre
// type controle : 0 controle + proposition du type
//                 1 controle + proposition du code
//                 2 controle strict
begin
    if TypeCodeBarre = '39C' then
    begin
         Controle39C (CodeBarre, ValeurPoss, TypeControle);
    end
{    else if TypeCodeBarre ='39' then
    begin
         Controle39 (CodeBarre);
    end }
    else if TypeCodeBarre ='128' then
    begin
         Controle128 (CodeBarre, ValeurPoss, TypeControle);
    end
    else if TypeCodeBarre ='E13' then
    begin
         ControleEAN (CodeBarre, ValeurPoss, TypeControle, 13);
    end
    else if TypeCodeBarre ='EA8' then
    begin
         ControleEAN (CodeBarre, ValeurPoss, TypeControle, 8);
    end
    else if TypeCodeBarre ='ITC' then
    begin
         ControleITC (CodeBarre);
    end
    else if TypeCodeBarre ='ITF' then
    begin
         ControleITF (CodeBarre);
    end
end;

end.
