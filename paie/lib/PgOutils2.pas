{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. : Stockage des procédures communes
Mots clefs ... : PAIE;OUTILS
*****************************************************************}
{
PT1   : 31/01/2005 SB V_60 Intégration de proc. communes
PT2   : 11/07/2005 PH V_60 FQ 11337 On accepte les cumuls alphanumériques
PT3   : 19/07/2005 SB V_65 Simplification de la function Collezerodevant
PT4   : 19/10/2005 VG V_60 Fonction SupprimeFichier déplacée dans PgOutils2
PT5   : 03/04/2006 VG V_65 Ajout fonction PaysISOLibEnvers
PT6   : 05/04/2006 VG V_65 Ajout fonction CalculCleSS
PT7   : 17/07/2006 VG V_70 Ajout fonction PGSupprimeaccent
PT8   : 04/09/2006 VG V_70 Fignolage contrôle du SIRET - FQ N°12076
PT9   : 02/10/2006 PH V_70 Mise en place des controles Paie 50
PT10  : 17/10/2006 VG V_70 Utilisation d'une procédure générique pour supprimer
                           les accents
PT10  : 22/05/2007 VG V_72 Déplacement de la fonction TestNumeroSSNaissance
PT11  : 10/10/2007 VG V_80 Correction NODOSSIER en multi PME
PT12  : 31/03/2008 FC V_90 FQ 15310 La civilité n'est pas mentionnée uniquement pour les hommes
}
unit PgOutils2;

interface

uses SysUtils,
     HEnt1,
     math,
{$IFDEF EAGLCLIENT}
     UTOB,
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     HCtrls,
     Controls,
     HMsgBox
     ;

procedure PgAffecteNoDossier ();
procedure ForceNumerique(Buf: string; var Resultat: string);
procedure MajEtabCompl;
function ControlSiret(NoSiret: string): Boolean;
function TestNumeroSS(numss: string; codesexe: string): integer;
function CalculCleSS(NumSS13 : string): integer;
function ColleZeroDevant(Nombre, LongChaine: integer): string;
function RendExerSocialPrec(var MoisE, AnneeE, ComboExer: string; var DebExer, FinExer: TDateTime; Exer: string): Boolean;
function AffectDefautCode(Edit: THedit; Long: integer): string;
function PGRendNoDossier(): string;
procedure ChargementTablette(Prefixe, Typ: string);
procedure InitialiseCombo(var Zone: Tcontrol);
function RendPeriodeEnCours(var ExerPerEncours, DebPer, FinPer: string): Boolean;
function RendExerSocialEnCours(var MoisE, AnneeE, ComboExer: string; var DebExer, FinExer: TDateTime): Boolean;
Function PgGetLibelleTablette(Var NomChamp,ValChamp,Tablette : string) : String;
function SupprimeFichier(FileN: string; Silence: Boolean = False): Boolean;
procedure PaysISOLib(Pays: string; var CodeIso, Libelle: string);
procedure PaysISOLibEnvers(Libelle: string; var CodeIso : string);
function PGUpperCase(const S: string; withAccent: boolean = true): string;
function PGSupprimeaccent(Ch: Char): Char;
function PGControlBL (LimitePaie : Integer) : Boolean ;
procedure TestNumeroSSNaissance (numss, SSNaiss : string; ChampDate : TDateTime;
                                 ChampDept : string; var ResultAnnee,
                                 ResultMois, ResultDepart: Integer);
var
cle: integer;


implementation
uses paramsoc;
{***********A.G.L.***********************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/07/2001
Modifié le ... :   /  /
Description .. : Récupération dans Result du numéro contenu dans Buf
Suite ........ : après épuration des caractères non numériques
Mots clefs ... : PAIE,CHAINE
*****************************************************************}
procedure ForceNumerique(Buf: string; var Resultat: string);
var
i, j: integer;
begin
j := 1;
Resultat := StringOfChar(' ', Length(Buf));

for i := 1 to Length(Buf) do
    if ((Buf[i] >= '0') and (Buf[i] <= '9')) then
       begin
       Resultat[j] := Buf[i];
       j := j + 1;
       end;

Resultat := Trim(Resultat);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure MajEtabCompl;
var
St, LeChamp: string;
Q: TQuery;
begin
St:= 'SELECT DH_NOMCHAMP'+
     ' FROM DECHAMPS WHERE'+
     ' DH_NOMCHAMP LIKE "ETB_%" AND'+
     ' ((DH_TYPECHAMP LIKE "VARC%") OR (DH_TYPECHAMP="COMBO"))';
Q := OpenSql(St, TRUE,-1,'',true);
while not Q.EOF do
      begin
      LeChamp := Q.FindField('DH_NOMCHAMP').AsString;
      st:= 'UPDATE ETABCOMPL SET '+
           LeChamp+' = "" WHERE '+
           LeChamp+' IS NULL';
      ExecuteSql(st);
      Q.Next;
      end;
Ferme(Q);
St:= 'SELECT DH_NOMCHAMP'+
     ' FROM DECHAMPS WHERE'+
     ' DH_NOMCHAMP LIKE "ETB_%" AND'+
     ' DH_TYPECHAMP="DATE"';
Q := OpenSql(St, TRUE,-1,'',true);
while not Q.EOF do
      begin
      LeChamp := Q.FindField('DH_NOMCHAMP').AsString;
      st:= 'UPDATE ETABCOMPL SET '+
           LeChamp+' = "'+UsDateTime(IDate1900)+'" WHERE '+
           LeChamp+' IS NULL';
      ExecuteSql(st);
      Q.Next;
      end;
Ferme(Q);
St:= 'SELECT DH_NOMCHAMP'+
     ' FROM DECHAMPS WHERE'+
     ' DH_NOMCHAMP LIKE "ETB_%" AND'+
     ' DH_TYPECHAMP="BOOLEAN"';
Q := OpenSql(St, TRUE,-1,'',true);
while not Q.EOF do
      begin
      LeChamp := Q.FindField('DH_NOMCHAMP').AsString;
      st:= 'UPDATE ETABCOMPL SET '+
           LeChamp+' = "-" WHERE '+
           LeChamp+' IS NULL';
      ExecuteSql(st);
      Q.Next;
      end;
Ferme(Q);
St:= 'SELECT DH_NOMCHAMP'+
     ' FROM DECHAMPS WHERE'+
     ' DH_NOMCHAMP LIKE "ETB_%" AND'+
     ' ((DH_TYPECHAMP="DOUBLE") OR (DH_TYPECHAMP="INTEGER"))';
Q := OpenSql(St, TRUE,-1,'',true);
while not Q.EOF do
      begin
      LeChamp := Q.FindField('DH_NOMCHAMP').AsString;
      st:= 'UPDATE ETABCOMPL SET '+
           LeChamp+' = 0 WHERE '+
           LeChamp+' IS NULL';
      ExecuteSql(st);
      Q.Next;
      end;
Ferme(Q);
end;


{***********A.G.L.***********************************************
Auteur  ...... : PAIE
Créé le ...... : 26/11/2001
Modifié le ... : 27/11/2001
Description .. :  fonction de contrôle de validité du N° Siret (ou SIREN)
Suite ........ : On multiplie les différents chiffres composant le n° par un
Suite ........ : coefficient 1 ou 2 alternativement en commeçant par le
Suite ........ : dernier chiffre du n° (le 14ème * 1, le 13ème * 2, le 12ème *
Suite ........ : 1...). Pour chaque produit obtenu on ajoute les deux chiffres
Suite ........ : du résultat. Le total des résultats doit être un nombre de
Suite ........ : dizaines se terminant par 0.
Suite ........ : En amont il faut contrôler si le Siret est renseigné et si la
Suite ........ : taille est correcte (14 pour le SIRET, 9 pour le SIREN)
Mots clefs ... : PAIE,SIRET
*****************************************************************}
function ControlSiret(NoSiret: string): Boolean;
var
coeff, posit, iposit, zcalc, wcalc1, wcalc2, totcalc : Integer;
acalc, scalc : string;
begin
coeff:= 1;
totcalc:= 0;
posit:= Length (NoSiret);
iposit:= posit;

//PT8
if (not (IsNumeric (NoSiret))) then
   begin
   result:= False;
   exit;
   end;
//FIN PT8

while iposit > 0 do
      begin
//PT8
      scalc:= Copy (NoSiret, iposit, 1);
      if ((scalc = '0') or (scalc = '1') or (scalc = '2') or (scalc = '3') or
         (scalc = '4') or (scalc = '5') or (scalc = '6') or (scalc = '7') or
         (scalc = '8') or (scalc = '9')) then
         zcalc:= StrToInt (scalc)
      else
         begin
         result:= False;
         exit;
         end;
//FIN PT8
      zcalc:= zcalc * coeff;
      if (zcalc < 10) then
         acalc:= '0'+IntToStr (zcalc)
      else
         acalc:= IntToStr (zcalc);

      wcalc1:= StrToInt (copy (acalc, 1, 1));
      wcalc2:= StrToInt (copy (acalc, 2, 1));
      totcalc:= totcalc+wcalc1+wcalc2;
      iposit:= iposit-1;
      coeff:= coeff+1;
      if (coeff > 2) then
         coeff:= 1;
      end;
if (copy (IntToStr (totcalc), 2, 1)) <> '0' then
   result:= False
else
   result:= True;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TestNumeroSS(numss: string; codesexe: string): integer;
var
code, i, k, numcle, resultat, temp : integer;
numss13, numsscle, numdep : string;
Corse: Boolean;
begin
k:= length(numss);
resultat:= 0;
Corse:= False;
if ((numss <> '') or (numss = '000000000000000')) then
   begin
   if (resultat = 0) and (k < 15) then
      resultat:= -3;
   if (resultat = -3) and (k = 13) then
      resultat:= -2;

//contrôle du code sexe
   if (k = 15) and (resultat = 0) then
      begin
      if ((numss[1] <> '1') and (numss[1] <> '2') and (numss[1] <> '7') and
         (numss[1] <> '8')) then
         resultat:= -4; //incohérence entre code sexe saisie et code sexe conforme
      if ((numss[1] = '1') or (numss[1] = '7')) then //psa_sexe = homme
         begin
         if (codesexe = 'F') then
            resultat:= -5 //Incohérence entre psa_sexe et code sexe
         else
            if (numss[1] = '7') then
               resultat:= 3;
         end
      else
         begin
         if ((numss[1] = '2') or (numss[1] = '8')) then //psa_sexe = femme
            begin
            if (codesexe = 'M') then
               resultat:= -6 //Incohérence entre psa_sexe et code sexe
            else
               if (numss[1] = '8') then
                  resultat:= 3;
            end;
         end;
      end;

//contrôle du code département Corse
   if (resultat >= 0) then
      begin
      if ((k = 15) and (numss[6] = '2') and
         ((numss[7] = 'A') or (numss[7] = 'B'))) then
         begin
         numdep:= copy(numss, 2, 2);
         if (isnumeric(numdep)) then
            if (strtoint(numdep)>=76) then
               Corse:= True;
         end;

//contrôle de numéricité du n° de sécu
      for i:= 1 to k do
          begin
          if (not((Corse = True) and (i=7))) then
             begin
             if not(numss[i] in ['0'..'9']) then
                resultat:= -7; //non numérique
             end;
          end;
      end;

//calcul de la clé
   if ((resultat >= 0) and (k = 15)) then
      begin
//Récupération 13 premier char et cle saisie
      numss13:= Copy(numss, 1, 13);
{PT6
      if (NumSS13[7] = 'A') or (NumSS13[7] = 'B') then
         NumSS13[7] := '0';
      if numss13 <> '' then
         begin
         Num13:= valeur(Numss13);
         numsscle:= numss[14] + numss[15];
         val(numsscle, temp, code);
         numcle:= temp;
         entier:= Int(Num13 / 97);
         DblCle:= 97 - (Num13 - (entier * 97));
         Cle:= StrToInt(FloatToStr(DblCle));
         if (cle <> numcle) and (numcle > 0) then
            resultat:= -1
         end;
}
      Cle:= CalculCleSS (numss13);
      numsscle:= numss[14] + numss[15];
      val(numsscle, temp, code);
      numcle:= temp;
      if (cle <> numcle) and (numcle > 0) then
         resultat:= -1
//FIN PT6
      end;
   end
else
   if numss = '' then
      resultat:= 1;

result := resultat;
end;


//PT6
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/04/2006
Modifié le ... :   /  /
Description .. : Fonction de calcul du numéro de clé SS à partir du numéro
Suite ........ : sur 13 caractères
Mots clefs ... :
*****************************************************************}
function CalculCleSS(NumSS13 : string): integer;
var
entier, num13 : Extended;
DblCle : Double;
begin
result:= 0;
if (NumSS13[7] = 'A') or (NumSS13[7] = 'B') then
   NumSS13[7] := '0';
if numss13 <> '' then
   begin
   Num13:= valeur(Numss13);
   entier:= Int(Num13 / 97);
   DblCle:= 97 - (Num13 - (entier * 97));
   result:= StrToInt(FloatToStr(DblCle));
   end;
end;
//FIN PT6

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function ColleZeroDevant(Nombre, LongChaine: integer): string;
begin
{ DEB PT3 }
if (Length(IntToStr(Nombre)) < LongChaine)  then
   Result:= StringOfChar('0',LongChaine-Length(IntToStr(Nombre)))+IntToStr(Nombre)
else
     Result := IntToStr(Nombre);
{ FIN PT3 }
{ PT3 Mise en commentaire
var  StNombre: string;
TabInt: string;
i, j: integer;

tabResult := '';
for i := 1 to LongChaine do
    begin
    if Nombre < power(10, i) then
       begin
       TabInt := inttostr(Nombre);
//       colle (LongChaine-i zéro devant]
       for j := 0 to (LongChaine - i - 1) do
           insert('0', TabResult, j);
       result := concat(TabResult, Tabint);
       exit;
       end;
    if i > LongChaine then
       result := inttostr(Nombre);
    end;
}
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function RendExerSocialPrec(var MoisE, AnneeE, ComboExer: string; var DebExer, FinExer: TDateTime; Exer: string): Boolean;
var
Q: TQuery;
DatF: TDateTime;
Jour, Mois, Annee: WORD;
begin
result := FALSE;
DebExer := 0;
FinExer := 0;
Q:= OpenSQL ('SELECT *'+
             ' FROM EXERSOCIAL WHERE'+
             ' PEX_ANNEEREFER="'+Exer+'" AND'+
             ' PEX_ACTIF="X"', TRUE,-1,'',true);
if not Q.EOF then
   begin
   DatF := Q.FindField('PEX_FINPERIODE').AsFloat; // Recup date de fin periode en cours
   DecodeDate(DatF, Annee, Mois, Jour);
   MoisE := ColleZeroDevant(Mois, 2);
   ComboExer := Q.FindField('PEX_EXERCICE').AsString; // recup Combo identifiant exercice
   AnneeE := Q.FindField('PEX_ANNEEREFER').AsString; // recup Annee de exercice
   DebExer := Q.FindField('PEX_DATEDEBUT').AsDateTime;
   FinExer := Q.FindField('PEX_DATEFIN').AsDateTime;
   result := TRUE;
   end;
Ferme(Q);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function AffectDefautCode(Edit: THedit; Long: integer): string;
begin
// DEB PT2
if Edit <> nil then
   if (edit.text <> '') AND (IsNumeric (edit.text)) then
      result := ColleZeroDevant(StrToInt(Edit.text), Long)
   else
      begin
      if NOT (IsNumeric (edit.text)) then
         result := edit.text
      else
         result := '';
      end;
// FIN PT2
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function PGRendNoDossier(): string;
begin
if V_PGI.NoDossier <> '' then
   result := V_PGI.NoDossier
else
   result := '000000';
end;

procedure PgAffecteNoDossier ();
Var LeDos : String;
begin
leDos:= GetParamSocSecur ('SO_NODOSSIER', '');
{PT11
if (LeDos<>'000000') AND (LeDos<>'') then
   V_PGI.NoDossier:= LeDos
else
if (LeDos='') then
   V_PGI.NoDossier:= '000000';
}
if (LeDos<>'000000') AND (LeDos<>'') then
   V_PGI.NoDossier:= LeDos
else
if ((LeDos='') and (Length (V_PGI.NoDossier)=8)) then
   V_PGI.NoDossier:= '000000';
//FIN PT11
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. : En validation et en suppression d'enregistrement
Suite ........ : on recharge les tablettes.
Suite ........ : Sur le prefixe de la table on recharge les tablettes concernés
Suite ........ : On peut recharger qu'une seule tablette en renseignant le type
Mots clefs ... :
*****************************************************************}
procedure ChargementTablette(Prefixe, Typ: string);
var
QQuery: TQuery;
StWhere, StPref, StTyp: string;
begin
StWhere := '';
StPref := '';
StTyp := '';
if Prefixe <> '' then
   StPref:= ' DO_PREFIXE="'+Prefixe+'" ';
if Typ <> '' then
   StTyp:= ' DO_TYPE="'+Typ+'" ';
if StPref <> '' then
   StWhere:= 'WHERE'+StPref;
if (StWhere = '') and (StTyp <> '') then
   StWhere:= 'WHERE'+StTyp;
if (StWhere <> '') and (StTyp <> '') then
   StWhere:= StWhere+'AND'+StTyp;

QQuery:= OpenSql('SELECT DO_COMBO FROM DECOMBOS '+StWhere+'', true,-1,'',true);
while not QQuery.EOF do
      begin
      Avertirtable(QQuery.Findfield('DO_COMBO').asstring);
      QQuery.Next;
      end;
Ferme(QQuery);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure InitialiseCombo(var Zone: Tcontrol);
var
Combo: THValCombobox;
begin
if zone = nil then exit;
Combo := THValcombobox(Zone);
if ((Combo <> nil) and (Combo.value = '')) then
   Combo.Itemindex := 0;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. : Fonction qui rend les periodes en cours du dernier exercice
Suite ........ : actif
Mots clefs ... : PAIE
*****************************************************************}
function RendPeriodeEnCours(var ExerPerEncours, DebPer, FinPer: string): Boolean;
var
Q: TQuery;
begin
result := FALSE;
DebPer := DateToStr(idate1900);
FinPer := DateToStr(idate1900);
Q:= OpenSQL ('SELECT *'+
             ' FROM EXERSOCIAL WHERE'+
             ' PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', TRUE,-1,'',true);
if not Q.EOF then
   begin
   DebPer := Q.FindField('PEX_DEBUTPERIODE').AsString;
   FinPer := Q.FindField('PEX_FINPERIODE').AsString;
   ExerPerEncours := Q.FindField('PEX_EXERCICE').AsString;
   result := TRUE;
   end;
Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. : Fonction qui rend l'année et le mois de l'exercice en cours
Suite ........ : Elle donne en fait le dernier exercice Actif
Mots clefs ... : PAIE
*****************************************************************}
function RendExerSocialEnCours(var MoisE, AnneeE, ComboExer: string; var DebExer, FinExer: TDateTime): Boolean;
var
Q: TQuery;
DatF: TDateTime;
Jour, Mois, Annee: WORD;
begin
result := FALSE;
DebExer := idate1900;
FinExer := idate1900;
Q:= OpenSQL ('SELECT *'+
             ' FROM EXERSOCIAL WHERE'+
             ' PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', TRUE,-1,'',true);
if not Q.EOF then
   begin
   DatF := Q.FindField('PEX_FINPERIODE').AsFloat;
   DecodeDate(DatF, Annee, Mois, Jour);
   MoisE := ColleZeroDevant(Mois, 2);
   ComboExer := Q.FindField('PEX_EXERCICE').AsString;
   AnneeE := Q.FindField('PEX_ANNEEREFER').AsString;
   DebExer := Q.FindField('PEX_DATEDEBUT').AsDateTime;
   FinExer := Q.FindField('PEX_DATEFIN').AsDateTime;
   result := TRUE;
   end;
Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 04/03/2005
Modifié le ... :   /  /
Description .. : Recherche le libellé d'une tablette par rapport au suffice du
Suite ........ : champ
Mots clefs ... : PAIE;
*****************************************************************}
Function PgGetLibelleTablette( Var NomChamp,ValChamp,Tablette : string) : String;
Var Q : TQuery;
    St : string;
Begin
Tablette := '';
Q:= OpenSql ('SELECT DO_COMBO,DO_NOMCHAMP'+
             ' FROM DECOMBOS WHERE'+
             ' (DO_NOMCHAMP LIKE "%'+NomChamp+'%" OR'+
             ' DO_COMBO = "'+NomChamp+'") AND'+
             ' (DO_COMBO LIKE "PG%" OR'+
             ' DO_COMBO LIKE "YY%" OR'+
             ' DO_COMBO LIKE "TT%")', True,-1,'',true);
while not Q.eof do
      begin
      if Q.FindField('DO_COMBO').AsString = NomChamp then
         begin //Si NomChamp=NomTablette
         Tablette := Q.FindField('DO_COMBO').asstring;
         Break;
         end;
      St := Q.FindField('DO_NOMCHAMP').AsString;
      while Trim(St) <> '' do //Si NomChamp=Clé suffixe
            begin
            if ReadTokenSt(St) = NomChamp then
               begin
               Tablette := Q.FindField('DO_COMBO').asstring;
               Break;
               end;
            end;
      if Tablette <> '' then
         break;
      Q.Next;
      end;
Ferme(Q);
//PT12 PB de tablette car pas les mêmes données (Type YCV au lieu de CIV). Tablette spécifique
//pour la paie avec MR pour Monsieur au lieu de M. Forcer la tablette YYCIVILITE
if NomChamp = 'CIVILITE' then Tablette := 'YYCIVILITE';
if (Tablette <> '') then
   Result:= RechDom(Tablette, ValChamp, False)
else
   result := '';
End;

{***********A.G.L.***********************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Supprime un fichier
Mots clefs ... : PAIE;FICHIER
*****************************************************************}
function SupprimeFichier(FileN: string; Silence: Boolean = False): Boolean;
begin
result := True;
if FileExists(FileN) then
   begin
   if not Silence then
      begin
      if PgiAsk ('Voulez-vous supprimer le fichier '+ExtractFileName(FileN)+'?',
         'Fichier existant :') = MrYes then
         begin
         Result := True;
         DeleteFile(PChar(FileN))
         end
      else
         Result := False;
      end
   else
      begin
      Result := True;
      DeleteFile(PChar(FileN))
      end;
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/07/2002
Modifié le ... :   /  /
Description .. : Retourne le code ISO d'un pays sur 2 car et son libellé à
Suite ........ : partir du code ISO sur 3 car
Mots clefs ... : PAYS;CODEISO2;PAIE;PGDADSU
*****************************************************************}
procedure PaysISOLib(Pays: string; var CodeIso, Libelle: string);
var
Q: Tquery;
begin
CodeIso := '';
Libelle := '';
Q:= OpenSql ('SELECT PY_CODEISO2, PY_LIBELLE'+
            ' FROM PAYS WHERE'+
            ' PY_PAYS="'+Pays+'"', True,-1,'',true);
if not Q.EOF then
   begin
   CodeIso := Q.FindField('PY_CODEISO2').asString;
   Libelle := Q.FindField('PY_LIBELLE').asString;
   end;
ferme(Q);
end;

//PT5
{***********A.G.L.***********************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/04/2006
Modifié le ... :   /  /
Description .. : Retourne le code ISO d'un pays sur 3 car à
Suite ........ : partir de son libellé. En cas de pays non trouvé,
Suite ........ : initialisation à 'FRA'
Mots clefs ... : PAYS;CODEISO2;PAIE;PGDADSU
*****************************************************************}
procedure PaysISOLibEnvers(Libelle: string; var CodeIso : string);
var
Q: Tquery;
begin
CodeIso := '';
Libelle := '';
Q:= OpenSql ('SELECT PY_PAYS'+
            ' FROM PAYS WHERE'+
            ' PY_LIBELLE="'+Libelle+'"', True,-1,'',true);
if not Q.EOF then
   CodeIso := Q.FindField('PY_PAYS').asString
else
   CodeIso := 'FRA';

ferme(Q);
end;
//FIN PT5

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 15/07/2002
Modifié le ... :   /  /
Description .. : Remplace les minuscules par des majuscules non
Suite ........ : accentuées
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function PGUpperCase(const S: string; withAccent: boolean = true): string;
var
Ch: Char;
L: Integer;
Source, Dest: PChar;
begin
L:= Length(S);
SetLength(Result, L);
Source:= Pointer(S);
Dest:= Pointer(Result);
while L <> 0 do
      begin
      Ch:= Source^;
      if withAccent then
{PT10
         begin
         if Ch in ['á', 'à', 'â', 'ã', 'ä', 'å', 'Á', 'À', 'Â', 'Ã', 'Ä', 'Å'] then
            Ch := 'a'
         else
         if Ch in ['ç', 'Ç'] then
            Ch := 'c'
         else
         if Ch in ['é', 'è', 'ê', 'ë', 'É', 'È', 'Ê', 'Ë'] then
            Ch := 'e'
         else
         if Ch in ['í', 'ì', 'î', 'ï', 'Í', 'Ì', 'Î', 'Ï'] then
            Ch := 'i'
         else
         if Ch in ['ñ', 'Ñ'] then
            Ch := 'n'
         else
         if Ch in ['ó', 'ò', 'ô', 'õ', 'ö', 'Ó', 'Ò', 'Ô', 'Õ', 'Ö'] then
            Ch := 'o'
         else
         if Ch in ['ú', 'ù', 'û', 'ü', 'Ú', 'Ù', 'Û', 'Ü'] then
            Ch := 'u'
         else
         if Ch in ['ÿ', 'Ÿ'] then
            Ch := 'y';
         end;
}
         Ch:= PGSupprimeaccent (Ch);
//FIN PT10
      if (Ch >= 'a') and (Ch <= 'z') then
         Dec (Ch, 32);
      Dest^:= Ch;
      Inc (Source);
      Inc (Dest);
      Dec (L);
      end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/07/2006
Modifié le ... :   /  /
Description .. : Remplace les caractères par des caractères non
Suite ........ : accentuées
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function PGSupprimeaccent(Ch: Char): Char;
begin
result:= Ch;
if Ch in ['á', 'à', 'â', 'ã', 'ä', 'å', 'æ'] then
   result:= 'a'
else
if Ch in ['Á', 'À', 'Â', 'Ã', 'Ä', 'Å', 'Æ'] then
   result:= 'A'
else
if Ch in ['ç'] then
   result:= 'c'
else
if Ch in ['Ç'] then
   result:= 'C'
else
if Ch in ['é', 'è', 'ê', 'ë'] then
   result:= 'e'
else
if Ch in ['É', 'È', 'Ê', 'Ë'] then
   result:= 'E'
else
if Ch in ['í', 'ì', 'î', 'ï'] then
   result:= 'i'
else
if Ch in ['Í', 'Ì', 'Î', 'Ï'] then
   result:= 'I'
else
if Ch in ['ñ'] then
   result:= 'n'
else
if Ch in ['Ñ'] then
   result:= 'N'
else
if Ch in ['ó', 'ò', 'ô', 'õ', 'ö'] then
   result:= 'o'
else
if Ch in ['Ó', 'Ò', 'Ô', 'Õ', 'Ö'] then
   result:= 'O'
else
if Ch in ['ú', 'ù', 'û', 'ü'] then
   result:= 'u'
else
if Ch in ['Ú', 'Ù', 'Û', 'Ü'] then
   result:= 'U'
else
if Ch in ['ÿ'] then
   result:= 'y'
else
if Ch in ['Ÿ'] then
   result:= 'Y';
end;

// DEB PT9
function PGControlBL (LimitePaie : Integer) : Boolean ;
var St : String;
    Q : TQuery;
    NbreS : Integer;
begin
  st := 'Select count (*) FROM SALARIES WHERE PSA_DATESORTIE <= "' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE > "' + UsDateTime(Now) + '"';
  Q := OpenSql(St, TRUE,-1,'',true);
  if not Q.EOF then NbreS := Q.Fields[0].AsInteger;
  ferme(Q);
  if NbreS > LimitePaie then result := FALSE
  else result := true;
end;
// FIN PT9

//PT10
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/05/2007
Modifié le ... :   /  /
Description .. : fonction de contrôle des numéro SS pour : année, mois et département de naissance
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TestNumeroSSNaissance (numss, SSNaiss : string; ChampDate : TDateTime;
                                 ChampDept : string; var ResultAnnee,
                                 ResultMois, ResultDepart: Integer);
var
aa, mm, jj: word;
DateNaissanceSS: TDateTime;
AnneeSS, DepartNaissanceSS, MoisSS, TestMoisSS, TestNumDepartSS: string;
begin
if (SSNaiss='Annee') then
   ResultAnnee:= 0
else
if (SSNaiss='Mois') then
   ResultMois:= 0
else
if (SSNaiss='Depart') then
   ResultDepart:= 0;

if (SSNaiss='Mois') or (SSNaiss='Annee') then
   begin
   TestMoisSS:= Copy (numss, 4, 2);
   if IsNumeric (TestMoisSS) then
      begin
      if (StrToInt (TestMoisSS)>=20) and (StrToInt (TestMoisSS)<=42) then
         begin
         resultMois:= 0;
         ResultAnnee:= 0;
         exit;
         end;
      if StrToInt (TestMoisSS)>=50 then
         begin
         resultMois:= 0;
         ResultAnnee:= 0;
         exit;
         end;
      end;
   end;

if (SSNaiss = 'Mois') then
   begin
   if (numss<>'') or (numss='000000000000000') then
      begin
      DateNaissanceSS:= ChampDate;
      if DateNaissanceSS<>idate1900 then
         begin
         DecodeDate (DateNaissanceSS, aa, mm, jj);
         if mm<10 then
            MoisSS:= '0'+IntToStr (mm)
         else
            MoisSS:= IntToStr (mm);
         if (NumSS[4]<>MoisSS[1]) or (NumSS[5]<>MoisSS[2]) then
            ResultMois:= -9;
         end;
      end;
   end;

if (SSNaiss='Annee') then
   begin
   if ((numss<>'') or (numss='000000000000000')) then
      begin
      DateNaissanceSS:= ChampDate;
      if DateNaissanceSS<>idate1900 then
         begin
         DecodeDate (DateNaissanceSS, aa, mm, jj);
         AnneeSS:= IntToStr (aa);
         if (NumSS[2]<>AnneeSS[3]) or (NumSS[3]<>AnneeSS[4]) then
            ResultAnnee:= -8;
         end;
      end;
   end;

if (SSNaiss='Depart') then
   begin
//Controle département de naissance 75 du numss <> departement de naissance
//si Annee de naissance >= 1964 et <= 1968 alors Ok
   if ((numss<>'') or (numss='000000000000000')) then
      begin
      TestNumDepartSS:= Copy (numss, 6, 2);
      if (TestNumDepartSS='  ') or (Copy (TestNumDepartSS, 1, 1)=' ') or
         (Copy (TestNumDepartSS, 2, 1)=' ') then
         Exit;
      if IsNumeric (TestNumDepartSS) then
         begin
         if (StrToInt (TestNumDepartSS)>=91) and
            (StrToInt (TestNumDepartSS)<=96) and (ChampDept='99') then
            exit;
         end;
      if IsNumeric (ChampDept) then
         if (StrToInt (ChampDept)>=91) and (StrToInt(ChampDept)>=95) then
            exit;
      DepartNaissanceSS:= ChampDept;
      DateNaissanceSS:= ChampDate;
      aa:= 0;
      if DateNaissanceSS <> idate1900 then
         DecodeDate (DateNaissanceSS, aa, mm, jj);
      if Length (DepartNaissanceSS) = 2 then
         begin
         if (NumSS[6]<>DepartNaissanceSS[1]) or
            (NumSS[7]<>DepartNaissanceSS[2]) then
            ResultDepart:= -10;
         if (Copy (NumSS, 6, 2)='75') and
            (Copy (DepartNaissanceSS, 1, 2)<>Copy (NumSS, 6, 2)) and
            (aa <= 1968) then
            ResultDepart:= 0;
         if (DateNaissanceSS >= StrToDate ('01/01/1976')) then //salariés nés en corse
            begin
            if ((Copy (NumSS, 6, 2)='2A') or (Copy (NumSS, 6, 2)='2B')) and
               (DepartNaissanceSS = '20') then
               ResultDepart:= 0;
            end
         else
            begin
            if (DateNaissanceSS<>IDate1900) and ((Copy (NumSS, 6, 2)='2A') or
               (Copy (NumSS, 6, 2)='2B')) then
               ResultDepart:= -12;
            if (Copy (NumSS, 6, 2)='20') and ((DepartNaissanceSS='2A') or
               (DepartNaissanceSS = '2B')) then
               ResultDepart:= 0;
            end;
         end;
      end;
   end;
end;
//FIN PT10

end.
