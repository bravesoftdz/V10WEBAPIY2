unit uLibApercu;
interface

uses classes
     ,stdctrls
     ,uTob                              
     ,uTobDebug
     ;
const LIGNEVIDE : string = '%%VIDE%%';

procedure ChargeTOBApercu(var Memo : TMemo ; Apercu : TOB ; Separateur : Char = ' ');
function  GetLigneApercu(var Memo : TMemo ; NumLigne : integer) : String;
function AddLigneApercu(var Memo : TMemo ; NumLigne : integer ; Ligne : String = '' ; LigneTOB : TOB = nil ; Separateur : Char = ' ') : integer;
procedure DelLigneApercu(var Memo : TMemo ; NumLigne : integer);
function UpdateLigneApercu(var Memo : TMemo ; NumLigne : integer ; Ligne : String = '' ; LigneTOB : TOB = nil ; Separateur : Char = ' ') : integer;
function AddChampLigneApercu(var Memo : TMemo ; NumLigne : integer ; NumChamp : integer ; Champ : string ; Separateur : Char = ' ') : integer;
procedure DelChampLigneApercu(var Memo : TMemo ; NumLigne : integer ; NumChamp : integer ; Separateur : Char = ' ');
function UpdateChampLigneApercu(var Memo : TMemo ; NumLigne : integer ; NumChamp : integer ; Champ : string ; Separateur : Char = ' ') : integer;

procedure SplitString(var ChampsList : TStringList ;Ligne : string ; Separateur : Char = ' ') ;
function ToString(Value: Variant): String;
function RecupLigneTOB(Ligne : TOB; Separateur : Char = ' ') : string;
implementation
uses Sysutils
     ,HCtrls;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /
Description .. : Charge le contenu de la TOB dans le memo
Suite ........ : Chaque fille representant une ligne
Suite ........ : Chaque champ est séparé par le separateur
Mots clefs ... :
*****************************************************************}
procedure ChargeTOBApercu(var Memo : TMemo ; Apercu : TOB ; Separateur : Char = ' ');
var i : integer;
begin
  // On vide le memo :
  Memo.Lines.Clear ;
  for i := 0 to Apercu.Detail.Count - 1 do
     Memo.Lines.Add(RecupLigneTOB(Apercu.Detail[i],Separateur));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /    
Description .. : Retourne la ligne du Memo
Mots clefs ... : 
*****************************************************************}
function  GetLigneApercu(var Memo : TMemo ; NumLigne : integer) : String;
begin
  result := LIGNEVIDE;
  if NumLigne > Memo.Lines.Count then
     Exit
  else if NumLigne < 1 then
     Exit
  else
     result := Memo.Lines[NumLigne - 1];
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... : 13/02/2007
Description .. : Ajoute une ligne au Memo soit via une String soit via une
Suite ........ : TOB. Les autres lignes sont descendues d'un cran :
Suite ........ : 1
Suite ........ : 2
Suite ........ : 3
Suite ........ :
Suite ........ : 1
Suite ........ : Nouvelle Ligne 2
Suite ........ : 2
Suite ........ : 3
Mots clefs ... :
*****************************************************************}
function AddLigneApercu(var Memo : TMemo ; NumLigne : integer ; Ligne : String = '' ; LigneTOB : TOB = nil ; Separateur : Char = ' ') : integer;
var
  i       : integer;
begin
  // Vérification du compteur
  while NumLigne > Memo.Lines.Count + 1 do
  begin
     // Ajout de champs vides
     Memo.Lines.Add('')
  end;
  if NumLigne < 1 then
     i := 1
  else
     i := NumLigne - 1;

  if not assigned(LigneTOB) then
     Memo.Lines.Insert(i,Ligne)
  else
     Memo.Lines.Insert(i,RecupLigneTOB(LigneTOB,Separateur));
  result := i + 1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... : 13/02/2007
Description .. : Supprime la ligne spécifié dans le Memo si existe.
Suite ........ : 1
Suite ........ : A_Supprimer
Suite ........ : 2
Suite ........ : 3
Suite ........ : 
Suite ........ : 1
Suite ........ : 2
Suite ........ : 3
Mots clefs ... : 
*****************************************************************}
procedure DelLigneApercu(var Memo : TMemo ; NumLigne : integer);
begin
  // Verification du compteur
  if NumLigne > Memo.Lines.Count then
     Exit // Pas de ligne à supprimer
  else if NumLigne < 1 then
     Exit // Pas de ligne à supprimer
  else
     Memo.Lines.Delete(NumLigne - 1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... : 16/02/2007
Description .. : Modifie la ligne passée en paramètre dans le Memo
Mots clefs ... : 
*****************************************************************}
function UpdateLigneApercu(var Memo : TMemo ; NumLigne : integer ; Ligne : String = '' ; LigneTOB : TOB = nil ; Separateur : Char = ' ') : integer;
var
  i : integer;
begin
  result:=NumLigne;
  // Verification du compteur
  while NumLigne > Memo.Lines.Count do
  begin
     // Ajout de ligne vide
     AddLigneApercu(Memo,Memo.Lines.Count + 1);
  end;
  if NumLigne < 1 then
     Exit // Pas de ligne à modifier
  else
     i := NumLigne - 1;
  if not assigned(LigneTOB) then
     Memo.Lines[i] := Ligne
  else
     Memo.Lines[i] := RecupLigneTOB(LigneTOB,Separateur);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... : 13/02/2007
Description .. : Ajoute un champ à une ligne. Les autres champs sont 
Suite ........ : decalé vers la droite
Suite ........ : 1 2 3
Suite ........ : 1 Nouveau 2 3
Mots clefs ... : 
*****************************************************************}
function AddChampLigneApercu(var Memo : TMemo ; NumLigne : integer ; NumChamp : integer ; Champ : string ; Separateur : Char = ' ') : integer;
var
  ligne   : string;
  temp    : string;
  champs  : TStringList;
  i       : integer;

  begin
  result := 1;
  if NumLigne > Memo.Lines.Count then
  begin
     // La ligne n'existe pas encore on la rajoute
     // et on pose le champ à la premiere position.
     for i := Memo.Lines.Count to (NumLigne - 2) do
     begin
        AddLigneApercu(Memo,i+1,'',nil,Separateur);
     end;
     AddLigneApercu(Memo,NumLigne,Champ,nil,Separateur);
     Exit;
  end;
  ligne   := GetLigneApercu(Memo,NumLigne);
  if ligne = LIGNEVIDE then
     AddLigneApercu(Memo,NumLigne,Champ,nil,Separateur);
  champs  := TStringList.Create;
  try
     SplitString(champs,ligne,Separateur);
     temp := '';
     if (Champs.Count = 0) then
     begin
        // La ligne est vide on pose le champ à la premiere place.
        UpdateLigneApercu(Memo,NumLigne,Champ,nil,Separateur);
        result := 1;
        Exit;
     end
     else if (NumChamp = 1) then
        temp := Champ + Separateur + Champs[0]
     else
        temp := Champs[0];
     i := 1;
     while (i < champs.Count ) do
     begin
        if (NumChamp - 1) = i then
           temp := temp + Separateur + Champ + Separateur + Champs[i]
        else
           temp := temp + Separateur + Champs[i];
        Inc(i);
     end;
     if (NumChamp > i) then
     begin
        temp := temp + Separateur + Champ;
        NumChamp := i + 1;
     end;
  finally
     champs.free;
  end;
  UpdateLigneApercu(Memo,NumLigne,temp,nil,Separateur);
  result := NumChamp;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /    
Description .. : Supprime un champ sur une ligne
Suite ........ : 
Suite ........ : 1 A_Supprimer 2 3
Suite ........ : 1 2 3
Mots clefs ... : 
*****************************************************************}
procedure DelChampLigneApercu(var Memo : TMemo ; NumLigne : integer ; NumChamp : integer ; Separateur : Char = ' ');
var
  ligne   : string;
  temp    : string;
  champs  : TStringList;
  i       : integer;
begin
  ligne   := GetLigneApercu(Memo,NumLigne);
  if ligne = LIGNEVIDE then Exit;
  champs  := TStringList.Create;
  try
     SplitString(champs,ligne,Separateur);
     temp := '';
     for i := 0 to Champs.Count - 1 do
     begin
        if (NumChamp - 1) <> i then
        begin
           if temp <> '' then
              temp := temp + Separateur + Champs[i]
           else
              temp := Champs[i];
        end;
     end;
  finally
     champs.free;
  end;
  UpdateLigneApercu(Memo,NumLigne,temp,nil,Separateur);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /    
Description .. : Modifie un champ sur une ligne
Mots clefs ... : 
*****************************************************************}
function UpdateChampLigneApercu(var Memo : TMemo ; NumLigne : integer ; NumChamp : integer ; Champ : string ; Separateur : Char = ' ') : integer;
var   
  ligne   : string;
  temp    : string;
  champs  : TStringList;
  i       : integer;
begin
  result  := NumLigne;
  ligne   := GetLigneApercu(Memo,NumLigne);
  if ligne = LIGNEVIDE then
  begin
     // Nouvelle ligne
     Result := 1;
     temp := '';
     for i := 1 to NumChamp - 1 do
     begin
        temp := temp + Separateur;
     end;
     temp := temp + Champ;
     AddLigneApercu(Memo,NumLigne,temp,nil,Separateur);
     Exit
  end;
  champs  := TStringList.Create;
  try
     SplitString(champs,ligne,Separateur);
     temp := '';
     if (NumChamp > champs.Count) then
     begin
         for i := champs.Count + 1 to NumChamp - 1 do
             AddChampLigneApercu(Memo,NumLigne,i,' ',Separateur);
         Result := AddChampLigneApercu(Memo,NumLigne,NumChamp,Champ,Separateur);
     end
     else
     begin
        if (NumChamp = 1) then
           temp := Champ
        else
           temp := Champs[0];
        i := 1;
        while (i < champs.Count) do
        begin
           if (NumChamp - 1) = i then
              temp := temp + Separateur + Champ
           else
              temp := temp + Separateur + Champs[i];
           Inc(i);
        end;  
        UpdateLigneApercu(Memo,NumLigne,temp,nil,Separateur);
     end;
  finally
     champs.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /
Description .. : Fonction qui transforme un variant en string
Mots clefs ... : 
*****************************************************************}
function ToString(Value: Variant): String;
begin
  case TVarData(Value).VType of
    varSmallInt,
    varInteger   : Result := IntToStr(Value);
    varSingle,
    varDouble,
    varCurrency  : Result := FloatToStr(Value);
    varDate      : Result := usdateTime(Value);
    varBoolean   : if Value then Result := 'X' else Result := '-';
    varString    : Result := Value;
    else           Result := '';
  end;
  if result = #0 then result := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /
Description .. : Retourne une string composé des champs de la TOB 
Suite ........ : separés par Separateur
Mots clefs ... :
*****************************************************************}
function RecupLigneTOB(Ligne : TOB; Separateur : Char = ' ') : string;
var                              
  j : integer;
  temp : string;
begin
  // Recupération des champs Rééls
  for j := 1 to Ligne.GetChampCount(ttcReal) do
  begin
     Temp := ToString(Ligne.GetValeur(j));
     if Temp = '' then Temp := Separateur;
     if result <> '' then
        result := result + Separateur + Temp
     else
        result := Temp;
  end;
  // Récupération des champs Virtuels
  for j := 1000 to ( Ligne.GetChampCount(ttcSup) + 999 ) do 
  begin
     Temp := ToString(Ligne.GetValeur(j));
     if Temp = '' then Temp := Separateur;
     if result <> '' then
        result := result + Separateur + Temp
     else
        result := Temp;
  end;
end;
          
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : Explose une string en TStringList à l'aide d'un séparateur
Mots clefs ... : 
*****************************************************************}
procedure SplitString(var ChampsList : TStringList ;Ligne : string ; Separateur : Char = ' ') ;
var
  i      : integer;
  temp   : string ;
begin
  if ligne = '' then exit;
  temp := ligne;
  i := Pos(Separateur,temp);
  while (i <> 0) do
  begin
     ChampsList.Add(Copy(temp,0,i-1));
     // Delete(temp,0,i);
     temp := copy(temp,i+1,length(temp)-i);
     i := Pos(Separateur,temp);
  end;
  ChampsList.Add(Copy(temp,0,length(temp)));
end;
end.


